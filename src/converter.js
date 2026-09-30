import JsBarcode from "jsbarcode";
import Papa from "papaparse";

const EXPECTED_COLUMN_COUNT = 12;
const MONTH_NAMES_SHORT = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

export const DEFAULT_PREFERENCES = Object.freeze({
  bulkyItemsFirst: true,
  bulkyKeywords: ["KETER", "STORE IT OUT MAX"],
  includeBulkPickNote: true,
  includeTradeRedeliveryNote: true,
  includeItems: false,
  addDeliveryBorders: false
});

export function normalisePreferences(value = {}) {
  const keywords = Array.isArray(value.bulkyKeywords)
    ? value.bulkyKeywords
    : String(value.bulkyKeywords ?? "").split(/\r?\n/);

  const cleanedKeywords = [...new Map(
    keywords
      .map((keyword) => String(keyword).trim())
      .filter(Boolean)
      .map((keyword) => [keyword.toLocaleUpperCase("en-GB"), keyword])
  ).values()];

  return {
    bulkyItemsFirst: value.bulkyItemsFirst !== false,
    bulkyKeywords: cleanedKeywords.length > 0
      ? cleanedKeywords
      : [...DEFAULT_PREFERENCES.bulkyKeywords],
    includeBulkPickNote: value.includeBulkPickNote !== false,
    includeTradeRedeliveryNote: value.includeTradeRedeliveryNote !== false,
    includeItems: value.includeItems === true,
    addDeliveryBorders: value.addDeliveryBorders === true
  };
}

export function parseManifest(csvText, preferences = DEFAULT_PREFERENCES) {
  const options = normalisePreferences(preferences);
  const parsed = Papa.parse(String(csvText ?? ""), {
    skipEmptyLines: false
  });

  const quoteError = parsed.errors.find((error) => error.type === "Quotes");
  if (quoteError) {
    throw new Error(`The CSV could not be read: ${quoteError.message}`);
  }

  const rows = parsed.data.map((row) =>
    (Array.isArray(row) ? row : [row]).map((value) => String(value ?? ""))
  );

  if (rows.length === 0 || rows.every(isBlankRow)) {
    throw new Error("The CSV file is empty.");
  }

  rows[0][0] = rows[0][0].replace(/^\uFEFF/, "");

  if (rows[0].length !== EXPECTED_COLUMN_COUNT) {
    throw new Error(
      `Unexpected source column count: ${rows[0].length}. Expected ${EXPECTED_COLUMN_COUNT}.`
    );
  }

  const sourceRows = rows.slice(1)
    .filter((row) => !isBlankRow(row))
    .map((row) => {
      const padded = [...row];
      while (padded.length < EXPECTED_COLUMN_COUNT) padded.push("");
      return padded;
    });

  const orders = combineRowsByOrder(sourceRows, options);
  const vehicleMap = new Map();

  for (const order of orders) {
    const vehicle = order.vehicle.trim() || "Unassigned vehicle";
    const key = vehicle.toLocaleUpperCase("en-GB");
    if (!vehicleMap.has(key)) vehicleMap.set(key, { name: vehicle, orders: [] });
    vehicleMap.get(key).orders.push(order);
  }

  const vehicles = [...vehicleMap.values()].map((vehicle) => ({
    ...vehicle,
    orders: options.bulkyItemsFirst
      ? stableBulkyFirst(vehicle.orders)
      : [...vehicle.orders]
  }));

  if (vehicles.length === 0) {
    throw new Error("No route vehicles were found in the manifest.");
  }

  return {
    sourceRowCount: sourceRows.length,
    orderCount: orders.length,
    vehicleCount: vehicles.length,
    vehicles
  };
}

export async function buildManifestWorkbook(
  manifest,
  inputFileName,
  preferences = DEFAULT_PREFERENCES,
  barcodeRenderer = createBarcodeDataUrl
) {
  const options = normalisePreferences(preferences);
  const { default: ExcelJS } = await import("exceljs");
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Manifest Converter Web";
  workbook.company = "Developed by Chris Ryan";
  workbook.created = new Date();
  workbook.modified = new Date();

  const usedSheetNames = new Set();

  for (const vehicleGroup of manifest.vehicles) {
    const sheetName = safeSheetName(vehicleGroup.name, usedSheetNames);
    const worksheet = workbook.addWorksheet(sheetName, {
      views: [{
        state: "frozen",
        xSplit: 1,
        ySplit: 4,
        topLeftCell: "B5",
        activeCell: "B5",
        showGridLines: false
      }],
      pageSetup: {
        orientation: "landscape",
        paperSize: 9,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0
      }
    });

    const headers = [
      "Order barcode",
      "Notes",
      "Order ID",
      "Customer (last stop)",
      "Total weight",
      "Total Items"
    ];
    if (options.includeItems) headers.push("Items");

    const manifestDate = vehicleGroup.orders.find((order) => order.scheduledDate)?.scheduledDate;
    const subtitleParts = [`${vehicleGroup.orders.length} orders`];
    if (manifestDate) subtitleParts.push(`Scheduled ${formatDisplayDate(manifestDate)}`);
    const generatedAt = generatedTimestampFromFilename(inputFileName);
    if (generatedAt) subtitleParts.push(`Generated ${generatedAt}`);

    worksheet.getCell("A1").value = `Load Manifest — ${vehicleGroup.name}`;
    worksheet.getCell("A2").value = `${subtitleParts.join(". ")}.`;
    worksheet.getRow(4).values = headers;

    for (let index = 0; index < vehicleGroup.orders.length; index += 1) {
      const order = vehicleGroup.orders[index];
      const rowNumber = index + 5;
      const rowValues = [
        null,
        order.notes,
        order.orderId,
        order.customer,
        order.totalWeight,
        order.totalQuantity
      ];
      if (options.includeItems) rowValues.push(order.items);
      worksheet.getRow(rowNumber).values = rowValues;

      if (order.orderId) {
        const barcode = await barcodeRenderer(order.orderId);
        const imageId = workbook.addImage({ base64: barcode, extension: "png" });
        worksheet.addImage(imageId, {
          tl: { col: 0.08, row: rowNumber - 0.84 },
          ext: { width: 285, height: 22 },
          editAs: "oneCell"
        });
      }
    }

    formatWorksheet(worksheet, vehicleGroup.orders, headers.length, options);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return {
    buffer,
    outputFileName: buildOutputFileName(inputFileName),
    workbook
  };
}

export function createBarcodeDataUrl(value) {
  if (typeof document === "undefined") {
    throw new Error("Barcode rendering requires a browser canvas.");
  }

  const canvas = document.createElement("canvas");
  JsBarcode(canvas, value, {
    format: "CODE128",
    width: 2,
    height: 22,
    displayValue: false,
    margin: 0,
    background: "#ffffff",
    lineColor: "#000000"
  });
  return canvas.toDataURL("image/png");
}

export function buildOutputFileName(inputFileName) {
  const base = String(inputFileName || "manifest.csv").replace(/\.csv$/i, "");
  return `${base}_converted.xlsx`;
}

function formatWorksheet(worksheet, orders, lastColumn, options) {
  const lastRow = Math.max(4, orders.length + 4);

  for (let row = 1; row <= lastRow; row += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      worksheet.getCell(row, column).font = { name: "Arial", size: 12 };
    }
  }

  worksheet.getCell("A1").font = { name: "Arial", size: 15, bold: true };
  worksheet.getCell("A1").alignment = { horizontal: "left", vertical: "middle" };
  worksheet.getCell("A2").font = { name: "Arial", size: 10, italic: true };
  worksheet.getCell("A2").alignment = { horizontal: "left", vertical: "middle" };

  for (let column = 1; column <= lastColumn; column += 1) {
    const cell = worksheet.getCell(4, column);
    cell.font = { name: "Arial", size: 12, bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      ...(column === 1 ? { left: { style: "thin", color: { argb: "FF000000" } } } : {}),
      ...(column === lastColumn ? { right: { style: "thin", color: { argb: "FF000000" } } } : {})
    };
  }

  for (let row = 5; row <= lastRow; row += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      const cell = worksheet.getCell(row, column);
      cell.alignment = {
        ...cell.alignment,
        vertical: "middle",
        ...(column === 7 && options.includeItems ? { wrapText: true } : {})
      };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };

      if (options.addDeliveryBorders && row < lastRow) {
        cell.border = {
          bottom: { style: "medium", color: { argb: "FF7F7F7F" } }
        };
      }
    }
  }

  if (lastRow >= 5) {
    for (let row = 5; row <= lastRow; row += 1) {
      worksheet.getCell(row, 3).numFmt = "@";
      worksheet.getCell(row, 5).numFmt = "#,##0.00";
      worksheet.getCell(row, 6).numFmt = "#,##0";
    }
  }

  [42, 70, 22, 24, 17, 15, 48].slice(0, lastColumn).forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  worksheet.getRow(1).height = 26;
  worksheet.getRow(2).height = 20;
  worksheet.getRow(4).height = 34;

  for (let row = 5; row <= lastRow; row += 1) {
    const order = orders[row - 5];
    const estimatedLines = options.includeItems
      ? order.items.split("\n").reduce(
        (total, line) => total + Math.max(1, Math.ceil(line.length / 48)),
        0
      )
      : 1;
    worksheet.getRow(row).height = options.includeItems
      ? Math.min(409, Math.max(32, estimatedLines * 20 + 8))
      : 32;
  }
}

function combineRowsByOrder(sourceRows, options) {
  const groups = new Map();

  for (const row of sourceRows) {
    const rawOrderId = row[3];
    const orderId = cleanOrderId(rawOrderId);
    const key = orderId.toLocaleUpperCase("en-GB");

    if (!groups.has(key)) {
      groups.set(key, {
        baseRow: row,
        orderId,
        bulkPick: false,
        tradeOrRedelivery: false,
        bulkyPick: false,
        lines: []
      });
    }

    const group = groups.get(key);
    if (isBulkPickOrder(rawOrderId)) group.bulkPick = true;
    if (isTradeOrRedeliveryOrder(rawOrderId)) group.tradeOrRedelivery = true;

    const productId = row[11] ?? "";
    const rawProductName = row[10] ?? "";
    if (isBulkyProduct(rawProductName, options.bulkyKeywords)) group.bulkyPick = true;

    const productPrefix = productId ? `${productId}-` : "";
    const productName = productPrefix && rawProductName.startsWith(productPrefix)
      ? rawProductName.slice(productPrefix.length)
      : rawProductName;

    group.lines.push({
      productName,
      productId,
      quantity: parseNumber(row[9]),
      weight: parseNumber(row[8])
    });
  }

  return [...groups.values()].map((group) => {
    const baseRow = group.baseRow;
    return {
      vehicle: displayVehicleName(baseRow[0]),
      scheduledDate: parseSourceDate(baseRow[1]),
      orderId: group.orderId,
      customer: baseRow[4],
      servicePlan: baseRow[5],
      totalWeight: sumNumbers(group.lines.map((line) => line.weight)),
      totalQuantity: sumNumbers(group.lines.map((line) => line.quantity)),
      items: group.lines
        .map((line) => line.productName)
        .filter((name) => name.trim())
        .join("\n"),
      productIds: group.lines.map((line) => line.productId).join("\n"),
      bulkyPick: group.bulkyPick,
      notes: buildOrderNotes(group.bulkPick, group.tradeOrRedelivery, options)
    };
  });
}

function cleanOrderId(value) {
  let prefix = String(value ?? "").trim().split("-", 1)[0];
  if (/^E/i.test(prefix)) prefix = prefix.slice(1);
  const match = prefix.match(/^\d+/);
  return match ? match[0] : prefix;
}

function displayVehicleName(value) {
  return String(value ?? "").trim().replace(/\bRetail\b/gi, "Hertz");
}

function isBulkPickOrder(value) {
  const prefix = String(value ?? "").trim().split("-", 1)[0];
  return /_1B$/i.test(prefix);
}

function isTradeOrRedeliveryOrder(value) {
  return /^E/i.test(String(value ?? "").trim());
}

function isBulkyProduct(productName, keywords) {
  const name = String(productName ?? "").toLocaleUpperCase("en-GB");
  return keywords.some((keyword) => name.includes(keyword.trim().toLocaleUpperCase("en-GB")));
}

function buildOrderNotes(bulkPick, tradeOrRedelivery, options) {
  const notes = [];
  if (bulkPick && options.includeBulkPickNote) notes.push("STH&BULK");
  if (tradeOrRedelivery && options.includeTradeRedeliveryNote) {
    notes.push("Trade/Redelivery - may not be in location");
  }
  return notes.join(" | ");
}

function parseNumber(value) {
  const text = String(value ?? "").trim().replaceAll(",", "");
  if (!text) return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function sumNumbers(values) {
  const numbers = values.filter((value) => Number.isFinite(value));
  return numbers.length > 0
    ? numbers.reduce((total, value) => total + value, 0)
    : null;
}

function parseSourceDate(value) {
  const match = String(value ?? "").trim().match(
    /^(\d{4})\/(\d{2})\/(\d{2}) (\d{2}):(\d{2}):(\d{2})$/
  );
  if (!match) return null;
  const [, year, month, day, hour, minute, second] = match.map(Number);
  const date = new Date(year, month - 1, day, hour, minute, second);
  return date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day
    ? date
    : null;
}

function generatedTimestampFromFilename(fileName) {
  const match = String(fileName ?? "").match(
    /_(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})\.csv$/i
  );
  if (!match) return "";
  const [, dayText, monthText, yearText, hour, minute] = match;
  const day = Number(dayText);
  const month = Number(monthText);
  const year = Number(yearText);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return "";
  }
  return `${day} ${MONTH_NAMES_SHORT[month - 1]} ${year} at ${hour}:${minute}`;
}

function formatDisplayDate(date) {
  return `${date.getDate()} ${MONTH_NAMES_SHORT[date.getMonth()]} ${date.getFullYear()}`;
}

function safeSheetName(value, usedNames) {
  let baseName = String(value ?? "").trim() || "Unassigned vehicle";
  baseName = baseName.replace(/[:\\/?*[\]]/g, "_").slice(0, 31) || "Vehicle";
  let candidate = baseName;
  let suffix = 2;

  while (usedNames.has(candidate.toLocaleUpperCase("en-GB"))) {
    const suffixText = ` (${suffix})`;
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }

  usedNames.add(candidate.toLocaleUpperCase("en-GB"));
  return candidate;
}

function stableBulkyFirst(orders) {
  return orders
    .map((order, index) => ({ order, index }))
    .sort((left, right) =>
      Number(right.order.bulkyPick) - Number(left.order.bulkyPick) ||
      left.index - right.index
    )
    .map(({ order }) => order);
}

function isBlankRow(row) {
  return row.every((value) => !String(value ?? "").trim());
}
