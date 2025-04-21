package com.prollpa.serviceImpl;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.prollpa.exception.ResourceNotFoundException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.util.*;

@Service
public class ExcelToJsonServiceImpl {
    private static final Logger logger = LoggerFactory.getLogger(ExcelToJsonServiceImpl.class);

    // Retrieve all sheet names from the Excel file
    public List<String> getAllSheet(MultipartFile file) {
        List<String> sheetNames = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException(e.getMessage());
        }
        return sheetNames;
    }

     //Convert Excel to generic JSON format
    public byte[] convertExcelToJson(MultipartFile file, String sheetName) {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new ResourceNotFoundException("Sheet not present");

            List<Map<String, String>> jsonData = new ArrayList<>();
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new ResourceNotFoundException("Header row not present");

            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) headers.add(cell.getStringCellValue());

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Map<String, String> rowData = new LinkedHashMap<>();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    rowData.put(headers.get(j), cell == null ? "" : getCellValue(cell));
                }
                jsonData.add(rowData);
            }

            return new ObjectMapper().writerWithDefaultPrettyPrinter().writeValueAsBytes(jsonData);
        } catch (Exception e) {
            throw new ResourceNotFoundException(e.getMessage());
        }
    }
    

//    public byte[] convertExcelToJson1(MultipartFile file, String sheetName, List<String> columnNames) {
//        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
//            Sheet sheet = workbook.getSheet(sheetName);
//            if (sheet == null) throw new ResourceNotFoundException("Sheet not present");
//
//            List<Map<String, String>> jsonData = new ArrayList<>();
//            Row headerRow = sheet.getRow(0);
//            if (headerRow == null) throw new ResourceNotFoundException("Header row not present");
//
//            // Default to columns A and B if no column names are provided
//            List<String> headers =  columnNames != null && !columnNames.isEmpty() ? columnNames :Arrays.asList("B", "C");//columnNames != null && !columnNames.isEmpty() ? columnNames :
//
//            // Process rows with data present in at least the first two columns
//            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//                Row row = sheet.getRow(i);
//                if (row == null) continue;
//
//                // Check if first two cells have data
//                if (isRowValid(row, 2)) {
//                    Map<String, String> rowData = new LinkedHashMap<>();
//                    for (String colName : headers) {
//                        int colIndex = getColumnIndex(colName);
//                        Cell cell = row.getCell(colIndex);
//                        rowData.put(colName, cell == null ? "" : getCellValue1(cell));
//                    }
//                    jsonData.add(rowData);
//                }
//            }
//
//            return new ObjectMapper().writerWithDefaultPrettyPrinter().writeValueAsBytes(jsonData);
//        } catch (Exception e) {
//            throw new ResourceNotFoundException(e.getMessage());
//        }
//    }
    public byte[] convertExcelToJson1(MultipartFile file, String sheetName, String labelName, String valueName) {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new ResourceNotFoundException("Sheet not present");

            Map<String, String> jsonData = new LinkedHashMap<>();

            int labelIndex = 1; // Default to Column B (Index 1)
            int valueIndex = 2; // Default to Column C (Index 2)

            Row headerRow = sheet.getRow(0); // Assuming first row might contain headers
            if (headerRow != null) {
                // Search for labelName and valueName in the first row
                for (Cell cell : headerRow) {
                    String cellValue = getCellValue1(cell).trim();
                    if (cellValue.equalsIgnoreCase(labelName)) {
                        labelIndex = cell.getColumnIndex();
                    } else if (cellValue.equalsIgnoreCase(valueName)) {
                        valueIndex = cell.getColumnIndex();
                    }
                }
            }

            // Process data from row 5 (index 4) to row 100 (index 99)
            int lastRowIndex = sheet.getLastRowNum(); 
            for (int i = 4; i <= lastRowIndex; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell keyCell = row.getCell(labelIndex);
                Cell valueCell = row.getCell(valueIndex);

                if (keyCell != null && valueCell != null) {
                    String key = getCellValue1(keyCell).trim();
                    String value = getCellValue1(valueCell).trim();

                    if (!key.isEmpty()) {
                        jsonData.put(key, value); // Store key-value pairs
                    }
                }
            }

            return new ObjectMapper().writerWithDefaultPrettyPrinter().writeValueAsBytes(jsonData);
        } catch (IOException e) {
            throw new ResourceNotFoundException("Error processing file: " + e.getMessage());
        }
    }

    // Check if at least `requiredCells` have data
    private boolean isRowValid(Row row, int requiredCells) {
        int filledCells = 0;
        for (int i = 0; i < requiredCells; i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                filledCells++;
            }
        }
        return filledCells >= requiredCells;
    }

    // Convert Excel column letter (A, B, C, ...) to index (0, 1, 2, ...)
    private int getColumnIndex(String column) {
        return column.toUpperCase().charAt(0) - 'A';
    }

    // Get cell value as a string
    private String getCellValue1(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }


    // Convert Excel to FAQ JSON format
    public byte[] convertToFAQJSON(MultipartFile file) {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("FAQ");
            logger.info("Processing sheet: {}", sheet.getSheetName());

            JSONObject faqJson = new JSONObject();
            faqJson.put("subtitle5", getCellValue(sheet, 42));
            faqJson.put("subtitle4", getCellValue(sheet, 43));
            faqJson.put("title", getCellValue(sheet, 44));
            faqJson.put("subtitle2", getCellValue(sheet, 45));
            faqJson.put("subtitle3", getCellValue(sheet, 46));
            faqJson.put("subtitle1", getCellValue(sheet, 47));

            faqJson.put("content1", extractContent(sheet, 4, 11));
            faqJson.put("content2", extractContent(sheet, 12, 24));
            faqJson.put("content3", extractContent(sheet, 25, 27));
            faqJson.put("content4", extractContent(sheet, 28, 29));
            faqJson.put("content5", extractContent(sheet, 30, 41));

            JSONObject finalJson = new JSONObject();
            finalJson.put("faq", faqJson);
            logger.info("FAQ JSON generated successfully");

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            outputStream.write(finalJson.toString(4).getBytes(StandardCharsets.UTF_8));
            return outputStream.toByteArray();
        } catch (Exception e) {
            logger.error("Error processing file: {}", e.getMessage(), e);
            throw new ResourceNotFoundException(e.getMessage());
        }
    }

    // Get single-cell value safely
    private String getCellValue(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(1);
            if (cell != null) {
                return getCellValue(cell);
            }
        }
        return "";
    }

    // Extract FAQ content from rowsf
    private JSONArray extractContent(Sheet sheet, int startRow, int endRow) {
        JSONArray contentArray = new JSONArray();
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(1);
                if (cell != null) {
                    String cellValue = getCellValue(cell);
                    if (!cellValue.isEmpty()) {
                        processFAQEntry(cellValue, contentArray);
                    }
                }
            }
        }
        return contentArray;
    }

    // Process FAQ question and answer
    private void processFAQEntry(String cellValue, JSONArray contentArray) {
        boolean isArabic = containsArabic(cellValue);
        String delimiter = isArabic ? "؟" : "\\?";
        String[] parts = cellValue.split(delimiter, 2);
        if (parts.length < 2) return;

        String question = parts[0].trim() + (isArabic ? "؟" : "?");
        String answer = parts[1].trim();

        answer = answer.replaceAll("\\*\\*(.*?)\\*\\*", "<b>$1</b>");
        String direction = isArabic ? "rtl" : "ltr";

        String description = "<div style='direction: " + direction + "; text-align: " + (isArabic ? "right" : "left") + ";'>"
                + "<div><p>" + answer.replace("\n", "<br>") + "</p></div></div>";

        JSONObject faqEntry = new JSONObject();
        faqEntry.put("heading", question);
        faqEntry.put("description", description);
        contentArray.put(faqEntry);
    }

    // Check if text contains Arabic characters
    private boolean containsArabic(String text) {
        return Pattern.compile("[\\u0600-\\u06FF]").matcher(text).find();
    }

    // Get cell value safely
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BLANK:
            default:
                return "";
        }
    }
    public byte[] generateICR(MultipartFile file, String sheetName, String labelName, String valueName) {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new ResourceNotFoundException("Sheet not present");
            
            StringBuilder icrContent = new StringBuilder();
            icrContent.append(valueName).append("\n");

            int labelIndex = 1; // Column B (Index 1)
            int valueIndex = 2; // Column C (Index 2)
            int startRow = 4; // Row index starts from 0, so B5 is index 4

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell labelCell = row.getCell(labelIndex);
                Cell valueCell = row.getCell(valueIndex);

                if (labelCell != null && valueCell != null) {
                    String label = labelCell.getStringCellValue().trim();
                    String value = "\"" + valueCell.toString().trim() + "\"";
                    icrContent.append(label).append("=").append(value).append(";\n");
                }
            }

            // Convert StringBuilder to byte array and return as a .txt file
            return icrContent.toString().getBytes(StandardCharsets.UTF_8);
        } catch (IOException e) {
            throw new ResourceNotFoundException("Error processing file"+ e);
        }
    }

}
