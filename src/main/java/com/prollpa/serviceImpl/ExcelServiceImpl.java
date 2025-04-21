package com.prollpa.serviceImpl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import java.io.*;
import java.util.*;

@Service
public class ExcelServiceImpl {
    @Autowired
    private ExcelToJsonServiceImpl excelService;
    private String outputFilePath = "src/main/resources/test.txt";
    
    private static final Set<String> ALLOWED_SHEETS = Set.of(
            "Website Menu", "Website Content", "Website Document Checklist", "Information and Guidelines", "ICR", "Appointment Letter", "Visa Center Page Output", "Vas Services", "Visa Center Page local language"
    );
    
    public Map<String, String> processExcelFile(MultipartFile file) throws Exception {
        Map<String, String> response = new HashMap<>();
        String outputDir = System.getProperty("user.dir") + File.separator + "json_output";
        File directory = new File(outputDir);
        if (!directory.exists()) {
            directory.mkdirs();
        }
        
        try (InputStream is = file.getInputStream(); Workbook workbook = new XSSFWorkbook(is)) {
            for (Sheet sheet : workbook) {
                String sheetName = sheet.getSheetName();
                if (!ALLOWED_SHEETS.contains(sheetName)) {
                    continue;
                }
                
                Map<String, Object> jsonMap = new LinkedHashMap<>();
                String jsonFilePath = outputDir + File.separator + sheetName.replaceAll("\\s+", "_") + ".json";
                
                if ("Website Menu".equalsIgnoreCase(sheetName) || "Appointment Letter".equalsIgnoreCase(sheetName)
                        || "Visa Center Page Output".equalsIgnoreCase(sheetName) || "Vas Services".equalsIgnoreCase(sheetName)
                        || "Visa Center Page local language".equalsIgnoreCase(sheetName)) {
                    jsonMap.putAll(processWebsiteMenu(sheet));
                } else if ("ICR".equalsIgnoreCase(sheetName)) {
                    generateICRFile(sheet, outputDir);
                    response.put("ICR", "ICR.txt file generated");
                } else if ("Website Document Checklist".equalsIgnoreCase(sheetName) || "Website Content".equalsIgnoreCase(sheetName)
                        ||"Information and Guidelines".equalsIgnoreCase(sheetName)) {
                    jsonMap.putAll(processHowToApplySheet(sheet));
                }
                
                if (!jsonMap.isEmpty()) {
                    ObjectMapper objectMapper = new ObjectMapper();
                    try (FileWriter jsonWriter = new FileWriter(jsonFilePath)) {
                        jsonWriter.write(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(jsonMap));
                    }
                    response.put(sheetName, "JSON file created");
                }
            }
        }
        return response;
    }
    
    private Map<String, Map<String, String>> processWebsiteMenu(Sheet sheet) {
        Map<String, Map<String, String>> headersMap = new LinkedHashMap<>();
        String currentHeader = null;
        
        for (int i = 4; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            
            Cell headerCell = row.getCell(0);
            Cell keyCell = row.getCell(1);
            Cell valueCell = row.getCell(2);
            
            if (headerCell != null && headerCell.getCellType() != CellType.BLANK) {
                currentHeader = getCellValueAsString(headerCell).replace("\n", "").replace(" ", "_");
                headersMap.putIfAbsent(currentHeader, new LinkedHashMap<>());
            }
            
            if (keyCell != null && valueCell != null && currentHeader != null) {
                headersMap.get(currentHeader).put(
                        getCellValueAsString(keyCell).replace("\n", "").replace(" ", "_"),
                        getCellValueAsString(valueCell)
                );
            }
        }
        return headersMap;
    }
    
    private void generateICRFile(Sheet sheet, String outputDir) throws IOException {
        File icrFile = new File(outputDir + File.separator + "ICR.txt");
        try (FileWriter writer = new FileWriter(icrFile)) {
            for (int i = 4; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                
                Cell labelCell = row.getCell(1);
                Cell valueCell = row.getCell(2);
                
                if (labelCell != null && valueCell != null) {
                    writer.write(labelCell.getStringCellValue().trim() + "= \"" + valueCell.toString().trim() + "\";\n");
                }
            }
        }
    }
    
    private Map<String, String> processHowToApplySheet(Sheet sheet) {
        Map<String, String> dataMap = new LinkedHashMap<>();
        for (int i = 4; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            
            Cell keyCell = row.getCell(0);
            Cell valueCell = row.getCell(1);
            
            if (keyCell != null && valueCell != null) {
                dataMap.put(getCellValueAsString(keyCell).replace("\t", "").replace(" ", "_"),
                        getCellValueAsString(valueCell).replace("\n", "<br>").replace("\n\n", "<br>").replace("\t", "").replace("Standard Mandatory Requirements", "<strong>Standard Mandatory Requirements:</strong><br>")
                        .replace("Additional Mission Specific Mandatory Requirements", "<br><strong>Additional Mission Specific Mandatory Requirements:</strong><br>")
                        .replace("Notes", "<br><strong>Notes:</strong><br>"));
            }
        }
        return dataMap;
    }
    
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        String cellValue = switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
        
        return cellValue.replaceAll("\\*\\*(.*?)\\*\\*", "<strong>$1</strong>");
    }
    public Map<String, String> processExcelFileIndividual(MultipartFile file,Set<String>jsonRequiredList) throws Exception {
    	
        Map<String, String> response = new HashMap<>();
        String outputDir = System.getProperty("user.dir") + File.separator + "json_output";
        File directory = new File(outputDir);
        if (!directory.exists()) {
            directory.mkdirs();
        }
        
        try (InputStream is = file.getInputStream(); Workbook workbook = new XSSFWorkbook(is)) {
            for (Sheet sheet : workbook) {
                String sheetName = sheet.getSheetName();
                if (!jsonRequiredList.contains(sheetName)) {
                    continue;
                }
                
                Map<String, Object> jsonMap = new LinkedHashMap<>();
                String jsonFilePath = outputDir + File.separator + sheetName.replaceAll("\\s+", "_") + ".json";
                
                if ("Website Menu".equalsIgnoreCase(sheetName) || "Appointment Letter".equalsIgnoreCase(sheetName)
                        || "Visa Center Page Output".equalsIgnoreCase(sheetName) || "Vas Services".equalsIgnoreCase(sheetName)
                        || "Visa Center Page local language".equalsIgnoreCase(sheetName)) {
                    jsonMap.putAll(processWebsiteMenu(sheet));
                } else if ("ICR".equalsIgnoreCase(sheetName)) {
                    generateICRFile(sheet, outputDir);
                    response.put("ICR", "ICR.txt file generated");
                } else if ("Website Document Checklist".equalsIgnoreCase(sheetName) || "Website Content".equalsIgnoreCase(sheetName)
                        ||"Information and Guidelines".equalsIgnoreCase(sheetName)) {
                    jsonMap.putAll(processHowToApplySheet(sheet));
                }
                
                if (!jsonMap.isEmpty()) {
                    ObjectMapper objectMapper = new ObjectMapper();
                    try (FileWriter jsonWriter = new FileWriter(jsonFilePath)) {
                        jsonWriter.write(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(jsonMap));
                    }
                    response.put(sheetName, "JSON file created");
                }
            }
        }
        return response;
    }
    
    
}
