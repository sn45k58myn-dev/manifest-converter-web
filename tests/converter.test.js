import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createCanvas } from "@napi-rs/canvas";
import ExcelJS from "exceljs";
import JsBarcode from "jsbarcode";
import { describe, expect, test } from "vitest";
import {
  DEFAULT_PREFERENCES,
  buildManifestWorkbook,
  buildOutputFileName,
  parseManifest
} from "../src/converter.js";

const currentDirectory = path.dirname(fileURLToPath(import.meta.url));
const fixturePath = path.resolve(
  currentDirectory,
  "../sample/Load_Manifest_21_09_2026_10_30.csv"
);

async function renderTestBarcode(value) {
  const canvas = createCanvas(300, 32);
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

async function fixtureText() {
  return fs.readFile(fixturePath, "utf8");
}

describe("manifest parsing", () => {
  test("preserves the proven default order, notes and totals", async () => {
    const manifest = parseManifest(await fixtureText(), DEFAULT_PREFERENCES);

    expect(manifest.sourceRowCount).toBe(6);
    expect(manifest.orderCount).toBe(5);
    expect(manifest.vehicleCount).toBe(2);

    const vehicleA = manifest.vehicles.find((vehicle) => vehicle.name === "Vehicle A");
    expect(vehicleA.orders.map((order) => order.orderId)).toEqual([
      "1060000002",
      "1060000004",
      "1060000001",
      "1060000003"
    ]);
    expect(vehicleA.orders[1].totalWeight).toBe(20);
    expect(vehicleA.orders[1].totalQuantity).toBe(2);
    expect(vehicleA.orders[1].items).toBe(
      "KETER STORE IT OUT MIDI 880L\nAccessory for Keter order"
    );
    expect(vehicleA.orders[3].notes).toBe(
      "STH&BULK | Trade/Redelivery - may not be in location"
    );
  });

  test("supports preferences and changes standalone Retail to Hertz", async () => {
    const csv = (await fixtureText()).replaceAll("Vehicle A", "B1 Retail");
    const manifest = parseManifest(csv, {
      bulkyItemsFirst: false,
      bulkyKeywords: ["KETER", "STORE IT OUT MAX"],
      includeBulkPickNote: false,
      includeTradeRedeliveryNote: false,
      includeItems: true,
      addDeliveryBorders: true
    });

    const hertz = manifest.vehicles.find((vehicle) => vehicle.name === "B1 Hertz");
    expect(hertz.orders.map((order) => order.orderId)).toEqual([
      "1060000001",
      "1060000002",
      "1060000003",
      "1060000004"
    ]);
    expect(hertz.orders.every((order) => order.notes === "")).toBe(true);
  });

  test("rejects files that do not use the 12-column manifest layout", () => {
    expect(() => parseManifest("A,B,C\n1,2,3", DEFAULT_PREFERENCES)).toThrow(
      "Expected 12"
    );
  });
});

describe("workbook generation", () => {
  test("creates the complete Items and delivery-border workbook", async () => {
    const preferences = {
      ...DEFAULT_PREFERENCES,
      includeItems: true,
      addDeliveryBorders: true
    };
    const manifest = parseManifest(await fixtureText(), preferences);
    const result = await buildManifestWorkbook(
      manifest,
      "Load_Manifest_21_09_2026_10_30.csv",
      preferences,
      renderTestBarcode
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(result.buffer);

    expect(result.outputFileName).toBe(
      "Load_Manifest_21_09_2026_10_30_converted.xlsx"
    );
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual([
      "Vehicle A",
      "Vehicle B"
    ]);

    const vehicleA = workbook.getWorksheet("Vehicle A");
    expect(vehicleA.getCell("A1").value).toBe("Load Manifest — Vehicle A");
    expect(vehicleA.getCell("A2").value).toBe(
      "4 orders. Scheduled 21 Sep 2026. Generated 21 Sep 2026 at 10:30."
    );
    expect(vehicleA.getCell("G4").value).toBe("Items");
    expect(vehicleA.getCell("G6").value).toBe(
      "KETER STORE IT OUT MIDI 880L\nAccessory for Keter order"
    );
    expect(vehicleA.getImages()).toHaveLength(4);
    expect(workbook.getWorksheet("Vehicle B").getImages()).toHaveLength(1);
    expect(vehicleA.pageSetup.orientation).toBe("landscape");
    expect(vehicleA.pageSetup.paperSize).toBe(9);
    expect(vehicleA.pageSetup.fitToWidth).toBe(1);
    expect(vehicleA.views[0].showGridLines).toBe(false);
    expect(vehicleA.views[0].state).toBe("frozen");
    expect(vehicleA.views[0].xSplit).toBe(1);
    expect(vehicleA.views[0].ySplit).toBe(4);

    for (const row of [5, 6, 7]) {
      for (let column = 1; column <= 7; column += 1) {
        const border = vehicleA.getCell(row, column).border;
        expect(border.bottom.style).toBe("medium");
        expect(border.bottom.color.argb).toBe("FF7F7F7F");
        expect(border.left).toBeUndefined();
        expect(border.right).toBeUndefined();
      }
    }

    for (let column = 1; column <= 7; column += 1) {
      expect(vehicleA.getCell(8, column).border.bottom).toBeUndefined();
    }

    await fs.mkdir(path.resolve(currentDirectory, "../test-output"), { recursive: true });
    await fs.writeFile(
      path.resolve(currentDirectory, "../test-output/web-converted-sample.xlsx"),
      result.buffer
    );
  });

  test("keeps the default six-column workbook clean", async () => {
    const manifest = parseManifest(await fixtureText(), DEFAULT_PREFERENCES);
    const result = await buildManifestWorkbook(
      manifest,
      "manifest.csv",
      DEFAULT_PREFERENCES,
      renderTestBarcode
    );
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(result.buffer);
    const vehicleA = workbook.getWorksheet("Vehicle A");

    expect(vehicleA.getCell("G4").value).toBeNull();
    expect(vehicleA.getCell("A5").border.bottom).toBeUndefined();
  });

  test("uses the established converted filename", () => {
    expect(buildOutputFileName("Example.CSV")).toBe("Example_converted.xlsx");
  });
});
