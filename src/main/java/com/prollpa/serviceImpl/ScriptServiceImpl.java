package com.prollpa.serviceImpl;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.prollpa.exception.ResourceNotFoundException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;


@Service
public class ScriptServiceImpl {
	private static final Logger logger = LoggerFactory.getLogger(ScriptServiceImpl.class);
	private static final String FILE_PATH = "src/main/resources/Rollout New Automation Template.xlsx";
	public ResponseEntity<Resource> generateScriptFile() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Scripts");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=countryVSCconfiguration.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	public ResponseEntity<Resource> getnationalitytranslationScript() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("New Lang Nationality UAT Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 250; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("Nationality Script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Nationality Script.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	public ResponseEntity<Resource> getLocalLanguageTranslationScript() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("NewLangCommon UATLive Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 250; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("NewLangCommon Script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=NewLangCommon Script.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	public ResponseEntity<Resource> generatecountryvisalinkUATOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("countryvisalink UAT Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 315; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=countryvisalink UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	public ResponseEntity<Resource> generateNewLangNationalityUATOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("New Lang Nationality UAT Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	
	public ResponseEntity<Resource> generateNewLangCommonUATLiveOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("NewLangCommon UATLive Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0); // Column A (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }

	
	
	
	public ResponseEntity<Resource> generateVisaCountryDocLinkOutpuLive() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Visa Country DocLink Outpu Live");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column A (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	
	public ResponseEntity<Resource> generateVisaCountryDocLinkOutputUAT() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Visa Country DocLink Output UAT");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }

	public ResponseEntity<Resource> generatecountryvisalinkLiveOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("countryvisalink Live Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	public ResponseEntity<Resource> generateVSCVisaInsuranceUATOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("VSC Visa Insurance UAT Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	public ResponseEntity<Resource> generateVSCVisaInsuranceLiveOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("VSC Visa Insurance Live Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	
	public ResponseEntity<Resource> generateCourierCityOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Courier City Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	public ResponseEntity<Resource> generateSMSforLocalLanguageOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("SMS for Local Language Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column B (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	public ResponseEntity<Resource> generateHolidayScriptOutput() throws InvalidFormatException {
        try {
            File excelFile = new File(FILE_PATH);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Holiday Script Output");

            if (sheet == null) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }

            StringBuilder scriptContent = new StringBuilder();
            for (int i = 4; i <= 300; i++) { // B5 to B20 (zero-based index)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(1); // Column A (index 1)
                    if (cell != null) {
                        scriptContent.append(cell.getStringCellValue().toString()).append("\n");
                    }
                }
            }
            workbook.close();

            // Write data to script.txt
            File scriptFile = new File("script.txt");
            try (FileWriter writer = new FileWriter(scriptFile)) {
                writer.write(scriptContent.toString());
            }

            // Return file for download
            Resource fileResource = new FileSystemResource(scriptFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=New Lang Nationality UAT Output.sql")
                    .body(fileResource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
	
	

}
