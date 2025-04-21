package com.prollpa.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class ExcelToTextFile {
    public static void main(String[] args) {
        // Define file paths
        String excelFilePath = "src/main/resources/privacy content automation.xlsx";
        String outputFilePath = "src/main/resources/test.txt";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath))) {

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            Sheet sheet = workbook.getSheet("Output");
            if (sheet == null) {
                System.out.println("Sheet 'Output' not found!");
                return;
            }

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0);

                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    continue;
                }

                if (cell.getCellType() == CellType.FORMULA) {
                    String formula = cell.getCellFormula();
                    System.out.println("Processing formula: " + formula);

                    // Updated regex to correctly match up to Function!G300
                    if (formula.matches("Function!G([1-9]|[1-9][0-9]|[1-2][0-9][0-9]|300)")) {
                        CellValue cellValue = formulaEvaluator.evaluate(cell);
                        String result = evaluateCellValue(cellValue);

                        if (!result.isEmpty() && !isZero(result)) {
                            writer.write(result.trim());
                            writer.newLine();
                        }
                    }
                } else {
                    String result = getCellValue(cell);
                    if (!result.isEmpty()) {
                        writer.write(result.trim());
                        writer.newLine();
                    }
                }
            }

            System.out.println("Process completed successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String evaluateCellValue(CellValue cellValue) {
        switch (cellValue.getCellType()) {
            case NUMERIC:
                return String.valueOf(cellValue.getNumberValue());
            case STRING:
                return cellValue.getStringValue();
            case BOOLEAN:
                return String.valueOf(cellValue.getBooleanValue());
            default:
                return "";
        }
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                double value = cell.getNumericCellValue();
                return value == 0 ? "" : String.valueOf(value);
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    private static boolean isZero(String result) {
        try {
            return Double.parseDouble(result) == 0;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}
