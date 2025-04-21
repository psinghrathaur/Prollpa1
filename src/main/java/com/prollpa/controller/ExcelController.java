package com.prollpa.controller;

import java.io.IOException;



import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.serviceImpl.ExcelServiceImpl;
import com.prollpa.serviceImpl.ExcelToJsonServiceImpl;
import com.prollpa.serviceImpl.PrivacyServiceImpl;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.springframework.web.bind.annotation.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


@RestController
@RequestMapping("/excel")
@Tag(name = "Excel Controller API", description = "Excel sheet apis")
public class ExcelController {
	private final PrivacyServiceImpl privacyServiceImpl = new PrivacyServiceImpl();
	private static final Logger logger = LoggerFactory.getLogger(ExcelController.class);
	@Autowired
	private ExcelToJsonServiceImpl excelService;
	@Autowired
	private ExcelServiceImpl excelRg;
	private ExcelToJsonServiceImpl excelToJson;
	
	@PostMapping("/sheets")
	@Operation(summary = "get the excel sheet names")
    public ResponseEntity<List<String>> getSheetNames(@RequestParam("file") MultipartFile file) {
        List<String> allSheet = excelService.getAllSheet(file);
        return ResponseEntity.ok(allSheet);
    }
	@PostMapping("/convertJson")
	@Operation(summary = "convertJson")
	public ResponseEntity<byte[]> convertExcelToJson(
	        @RequestParam("file") MultipartFile file,
	        @RequestParam(value = "sheetName", required = false, defaultValue = "")String sheetName,
	        @RequestParam(value="labelName",required=false,defaultValue="B")String labelName,
	        @RequestParam(value="valueName",required=false,defaultValue="C")String valueName) {
		String contentType = file.getContentType();
		
	    if (contentType == null || (!contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") 
	            && !contentType.equals("application/vnd.ms-excel"))) {
	        throw new ResourceNotFoundException("Invalid file type. Please upload an Excel file (.xls or .xlsx).");
	    }
	    // File Validation
	    if (file == null || file.isEmpty()) {
	        logger.warn("Uploaded file is null or empty");
	        return ResponseEntity.badRequest()
	                .body("Uploaded file is null or empty".getBytes());
	    }

	    if (!file.getOriginalFilename().endsWith(".xlsx")) {
	        logger.warn("Uploaded file is not an Excel file (.xlsx)");
	        return ResponseEntity.badRequest()
	                .body("Only .xlsx files are supported".getBytes());
	    }

	    try {
	        if (sheetName.isEmpty()) {
	            List<String> allSheet = excelService.getAllSheet(file);
	            if (allSheet.isEmpty()) {
	                logger.warn("The uploaded file does not contain any sheets");
	                return ResponseEntity.badRequest()
	                        .body("The uploaded file does not contain any sheets".getBytes());
	            }
	            sheetName = allSheet.get(0); // Default to the first sheet
	        }

	        logger.info("{} -> get", file.getOriginalFilename());

	        // Convert Excel to JSON
	        byte[] jsonBytes = excelService.convertExcelToJson1(file, sheetName,labelName,valueName);

	        // Prepare response headers
	        HttpHeaders headers = new HttpHeaders();
	        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=data.json");
	        headers.add(HttpHeaders.CONTENT_TYPE, "application/json");

	        return ResponseEntity.ok()
	                .headers(headers)
	                .body(jsonBytes);

	    } catch (Exception e) {
	        throw new ResourceNotFoundException(e.getMessage());
	    }
	}

	@PostMapping("/convertJson1")
	@Operation(summary = "convertJson1")
	public ResponseEntity<byte[]> convertExcelToJson1(
	        @RequestParam("file") MultipartFile file,
	        @RequestParam(value = "sheetName", required = false, defaultValue = "") String sheetName) {
		String contentType = file.getContentType();
		
	    if (contentType == null || (!contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") 
	            && !contentType.equals("application/vnd.ms-excel"))) {
	        throw new ResourceNotFoundException("Invalid file type. Please upload an Excel file (.xls or .xlsx).");
	    }
	    // File Validation
	    if (file == null || file.isEmpty()) {
	        logger.warn("Uploaded file is null or empty");
	        return ResponseEntity.badRequest()
	                .body("Uploaded file is null or empty".getBytes());
	    }

	    if (!file.getOriginalFilename().endsWith(".xlsx")) {
	        logger.warn("Uploaded file is not an Excel file (.xlsx)");
	        return ResponseEntity.badRequest()
	                .body("Only .xlsx files are supported".getBytes());
	    }

	    try {
	        if (sheetName.isEmpty()) {
	            List<String> allSheet = excelService.getAllSheet(file);
	            if (allSheet.isEmpty()) {
	                logger.warn("The uploaded file does not contain any sheets");
	                return ResponseEntity.badRequest()
	                        .body("The uploaded file does not contain any sheets".getBytes());
	            }
	            sheetName = allSheet.get(0); // Default to the first sheet
	        }

	        logger.info("{} -> get", file.getOriginalFilename());

	        // Convert Excel to JSON
	        byte[] jsonBytes = excelService.convertExcelToJson(file, sheetName);

	        // Prepare response headers
	        HttpHeaders headers = new HttpHeaders();
	        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=data.json");
	        headers.add(HttpHeaders.CONTENT_TYPE, "application/json");

	        return ResponseEntity.ok()
	                .headers(headers)
	                .body(jsonBytes);

	    } catch (Exception e) {
	        throw new ResourceNotFoundException(e.getMessage());
	    }
	}

	@PostMapping("/convertToFAQJSON")
	public ResponseEntity<byte[]> convertToFAQJSON(@RequestParam("file") MultipartFile file){
		 logger.info("Received request to generate downloadable FAQ JSON file");

	        if (file == null || file.isEmpty()) {
	            return ResponseEntity.badRequest().body(null);
	        }

	        if (!file.getOriginalFilename().endsWith(".xlsx")) {
	            return ResponseEntity.badRequest().body(null);
	        }

	        try {
	            byte[] jsonBytes = excelService.convertToFAQJSON(file);

	            HttpHeaders headers = new HttpHeaders();
	            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=faq.json");
	            headers.add(HttpHeaders.CONTENT_TYPE, "application/json");

	            return ResponseEntity.ok()
	                    .headers(headers)
	                    .body(jsonBytes);
	        } catch (Exception e) {
	            logger.error("Error processing file: {}", e.getMessage(), e);
	            throw new ResourceNotFoundException(e.getMessage());
	        }
		
	}
	
	@PostMapping("/generateICR")
	public ResponseEntity<byte[]> generateICR(@RequestParam("file") MultipartFile file,
	        @RequestParam(value = "sheetName", required = false, defaultValue = "") String sheetName,
	        @RequestParam(value="labelName",required=false,defaultValue="B")String labelName,
	        @RequestParam(value="valueName",required=false,defaultValue="C")String valueName){
		 logger.info("Received request to generate downloadable FAQ JSON file");

	        if (file == null || file.isEmpty()) {
	            return ResponseEntity.badRequest().body(null);
	        }

	        if (!file.getOriginalFilename().endsWith(".xlsx")) {
	            return ResponseEntity.badRequest().body(null);
	        }

	        try {
	            byte[] jsonBytes = excelService.generateICR(file,sheetName,labelName,valueName);

	            HttpHeaders headers = new HttpHeaders();
	            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ICR.txt");
	            headers.add(HttpHeaders.CONTENT_TYPE, "application/txt");

	            return ResponseEntity.ok()
	                    .headers(headers)
	                    .body(jsonBytes);
	        } catch (Exception e) {
	            logger.error("Error processing file: {}", e.getMessage(), e);
	            throw new ResourceNotFoundException(e.getMessage());
	        }
		
	}

	 @PostMapping("/upload")
	    public ResponseEntity<byte[]> uploadExcel(@RequestParam("file") MultipartFile file) {
	        try {
	            // Define output directory
	            String outputDir = System.getProperty("user.dir") + File.separator + "json_output";
	            File dir = new File(outputDir);
	            downloadPrivacyPolicy(file);
	            File privacyFile = new File("src/main/resources/test.txt");

	            // Ensure the directory exists
	            if (!dir.exists()) {
	                dir.mkdirs();
	            }

	            // Delete old JSON and ICR files before processing new ones
	            for (File fileEntry : dir.listFiles()) {
	                if (fileEntry.getName().endsWith(".json") || fileEntry.getName().equals("ICR.txt")) {
	                    fileEntry.delete();
	                }
	            }

	            // Process the Excel file (assumed to generate JSON and ICR files)
	            Map<String, String> response = excelRg.processExcelFile(file);

	            // Create ZIP file
	            File zipFile = new File(outputDir + File.separator + "output.zip");
	            try (FileOutputStream fos = new FileOutputStream(zipFile);
	                 ZipOutputStream zipOut = new ZipOutputStream(fos)) {

	                for (File fileEntry : dir.listFiles()) {
	                    if (!fileEntry.getName().endsWith(".json") && !fileEntry.getName().equals("ICR.txt")) {
	                        continue;
	                    }
	                    try (FileInputStream fis = new FileInputStream(fileEntry)) {
	                        ZipEntry zipEntry = new ZipEntry(fileEntry.getName());
	                        zipOut.putNextEntry(zipEntry);
	                        byte[] bytes = new byte[1024];
	                        int length;
	                        while ((length = fis.read(bytes)) >= 0) {
	                            zipOut.write(bytes, 0, length);
	                        }
	                        zipOut.closeEntry();
	                    }
	                }
	                if (privacyFile.exists()) {
	                    try (FileInputStream fis = new FileInputStream(privacyFile)) {
	                        ZipEntry zipEntry = new ZipEntry("privacy_policy.json");
	                        zipOut.putNextEntry(zipEntry);
	                        byte[] bytes = new byte[1024];
	                        int length;
	                        while ((length = fis.read(bytes)) >= 0) {
	                            zipOut.write(bytes, 0, length);
	                        }
	                        zipOut.closeEntry();
	                    }
	                }
	                byte[] convertToFAQJSON = excelService.convertToFAQJSON(file);
	                ZipEntry zipEntry = new ZipEntry("FAQ.json");
	                zipOut.putNextEntry(zipEntry);
	                zipOut.write(convertToFAQJSON, 0, convertToFAQJSON.length);
	                zipOut.closeEntry();
                    
                
	            }

	            // Read ZIP file and return as response
	            byte[] zipBytes = Files.readAllBytes(zipFile.toPath());
	            HttpHeaders headers = new HttpHeaders();
	            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=output.zip");
	            headers.add(HttpHeaders.CONTENT_TYPE, "application/zip");

	            return new ResponseEntity<>(zipBytes, headers, HttpStatus.OK);

	        } catch (Exception e) {
	            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
	                    .body(("Error processing file: " + e.getMessage()).getBytes());
	        }
	    
	}    
	    @PostMapping("/uploadIndividual")
	    public ResponseEntity<byte[]> uploadExcel(@RequestParam("file") MultipartFile file, 
	            @RequestParam("jsonRequiredList") String jsonRequiredList1) throws JsonMappingException, JsonProcessingException {
	        
	        System.out.println(file.getName() + " -> " + jsonRequiredList1);
	        
	        ObjectMapper objectMapper = new ObjectMapper();
	        Set<String> jsonRequiredList = objectMapper.readValue(jsonRequiredList1, new TypeReference<Set<String>>() {});
	        System.out.println(jsonRequiredList.contains("ICR"));
	        
	        File privacyFile = new File("src/main/resources/test.txt");

	        try {
	            // Define output directory
	            String outputDir = System.getProperty("user.dir") + File.separator + "json_output";
	            File dir = new File(outputDir);

	            // Ensure the directory exists
	            if (!dir.exists() && !dir.mkdirs()) {
	                throw new IOException("Failed to create output directory.");
	            }

	            // Define ZIP file path
	            File zipFile = new File(outputDir + File.separator + "output.zip");

	            // Delete old JSON, ICR, and ZIP files before processing new ones
	            File[] files = dir.listFiles();
	            if (files != null) {
	                for (File fileEntry : files) {
	                    if (fileEntry.getName().endsWith(".json") || fileEntry.getName().equals("ICR.txt") || fileEntry.getName().equals("output.zip")) {
	                        fileEntry.delete();
	                    }
	                }
	            }

	            // Process the Excel file (assumed to generate JSON and ICR files)
	            Map<String, String> response = excelRg.processExcelFileIndividual(file, jsonRequiredList);

	            // Create ZIP file
	            try (FileOutputStream fos = new FileOutputStream(zipFile);
	                 ZipOutputStream zipOut = new ZipOutputStream(fos)) {

	                files = dir.listFiles();
	                if (files != null) {
	                    for (File fileEntry : files) {
	                        // **Skip adding the output.zip file itself**
	                        if (fileEntry.getName().equals("output.zip")) {
	                            continue;
	                        }

	                        try (FileInputStream fis = new FileInputStream(fileEntry)) {
	                            ZipEntry zipEntry = new ZipEntry(fileEntry.getName());
	                            zipOut.putNextEntry(zipEntry);
	                            byte[] bytes = new byte[1024];
	                            int length;
	                            while ((length = fis.read(bytes)) >= 0) {
	                                zipOut.write(bytes, 0, length);
	                            }
	                            zipOut.closeEntry();
	                        }
	                    }
	                }

	                // Add Privacy Policy file if requested
	                if (jsonRequiredList.contains("Privacy Policy") && privacyFile.exists()) {
	                    try (FileInputStream fis = new FileInputStream(privacyFile)) {
	                        ZipEntry zipEntry = new ZipEntry("privacy_policy.json");
	                        zipOut.putNextEntry(zipEntry);
	                        byte[] bytes = new byte[1024];
	                        int length;
	                        while ((length = fis.read(bytes)) >= 0) {
	                            zipOut.write(bytes, 0, length);
	                        }
	                        zipOut.closeEntry();
	                    }
	                }

	                // Add FAQ JSON if requested
	                if (jsonRequiredList.contains("FAQ")) {
	                    byte[] convertToFAQJSON = excelService.convertToFAQJSON(file);
	                    ZipEntry zipEntry = new ZipEntry("FAQ.json");
	                    zipOut.putNextEntry(zipEntry);
	                    zipOut.write(convertToFAQJSON, 0, convertToFAQJSON.length);
	                    zipOut.closeEntry();
	                }

	            }

	            // Read ZIP file and return as response
	            byte[] zipBytes = Files.readAllBytes(zipFile.toPath());
	            HttpHeaders headers = new HttpHeaders();
	            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=output.zip");
	            headers.add(HttpHeaders.CONTENT_TYPE, "application/zip");

	            return ResponseEntity.ok().headers(headers).body(zipBytes);

	        } catch (Exception e) {
	            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
	                    .body(("Error processing file: " + e.getMessage()).getBytes(StandardCharsets.UTF_8));
	        }
	    }

	    
	    @PostMapping("/privacyPolicy")
	    public ResponseEntity<byte[]> downloadPrivacyPolicy(@RequestParam("file") MultipartFile file) {
	    	
	    	String status = privacyServiceImpl.writeDataFromUploadedExcel(file);
	    	System.out.println(status+"file");
	    	if(status!=null) {
	    	privacyServiceImpl.generateDataforPrivacyPolicy();
	        byte[] fileData = privacyServiceImpl.returnPrivacyPolicy();
	        System.out.println("generating file");
	        HttpHeaders headers = new HttpHeaders();
	        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Privacypolicy.json");
	        headers.add(HttpHeaders.CONTENT_TYPE, "text/plain");

	        return ResponseEntity.ok()
	                .headers(headers)
	                .body(fileData);
	    	}else {
	    		return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
		                
		                
	    	}
	    }
     
	

}
