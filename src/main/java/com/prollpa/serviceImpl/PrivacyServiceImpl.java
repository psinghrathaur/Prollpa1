package com.prollpa.serviceImpl;

import java.io.*;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.prollpa.exception.ResourceNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@Service
public class PrivacyServiceImpl {
	private static final int SOURCE_START_ROW = 4; // Row 4 in "Privacy Policy" (0-based index)
	private static final int SOURCE_END_ROW = 206; // Row 207 in "Privacy Policy"
	private static final int DEST_START_ROW = 2; // Row 3 in "Content"
	private static final int DEST_END_ROW = 205; // Row 206 in "Content"
	private static final int COLUMN_C = 2; // Column C (0-based index)
	private String excelFilePath = "src/main/resources/privacy content automation.xlsx";
    private final String outputFilePath = "src/main/resources/test.txt";
    public void generateDataforPrivacyPolicy() {
    	System.out.println("File generating start ");
    	resetFile();
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
                //System.out.println("Processing formula: " + formula);

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
 
	 
 public byte[] returnPrivacyPolicy() {
     try {
         Path path = Paths.get(outputFilePath);
         File file = path.toFile();

         if (!file.exists()) {
             throw new RuntimeException("File not found: " + outputFilePath);
         }

         return Files.readAllBytes(path);
     } catch (IOException e) {
         throw new RuntimeException("Error reading file: " + e.getMessage());
     }
 }
 
 public String writeDataFromUploadedExcel(MultipartFile file) {
	    try {
	        // Read the uploaded Excel file
	        Workbook sourceWorkbook = new XSSFWorkbook(file.getInputStream());
	        Sheet sourceSheet = sourceWorkbook.getSheet("Privacy Policy"); // Read from "Privacy Policy" sheet

	        if (sourceSheet == null) {
	            sourceWorkbook.close();
	            return "Uploaded file does not contain a sheet named 'Privacy Policy'!";
	        }

	        // Open the destination file
	        File destFile = new File(excelFilePath);
	        FileInputStream fis = new FileInputStream(destFile);
	        Workbook destWorkbook = new XSSFWorkbook(fis);
	        Sheet destSheet = destWorkbook.getSheet("Content"); // Target sheet: "Content"

	        if (destSheet == null) {
	            sourceWorkbook.close();
	            destWorkbook.close();
	            throw new ResourceNotFoundException("Destination sheet 'Content' not found!");
	            
	        }

	        // Copy data from Column C (Row 4-207) in "Privacy Policy" to Column C (Row 3-206) in "Content"
	        for (int srcRowNum = SOURCE_START_ROW, destRowNum = DEST_START_ROW; 
	             srcRowNum <= SOURCE_END_ROW && destRowNum <= DEST_END_ROW; 
	             srcRowNum++, destRowNum++) {

	            Row sourceRow = sourceSheet.getRow(srcRowNum);
	            Row destRow = destSheet.getRow(destRowNum);
	            if (destRow == null) destRow = destSheet.createRow(destRowNum);

	            if (sourceRow != null) {
	                Cell sourceCell = sourceRow.getCell(COLUMN_C); // Read from Column C in "Privacy Policy"
	                if (sourceCell != null) {
	                    Cell destCell = destRow.getCell(COLUMN_C);
	                    if (destCell == null) destCell = destRow.createCell(COLUMN_C);
	                    destCell.setCellValue(sourceCell.toString()); // Write into Column C in "Content"
	                }
	            }
	        }

	        // Save the updated file
	        fis.close();
	        FileOutputStream fos = new FileOutputStream(destFile);
	        destWorkbook.write(fos);
	        fos.close();

	        // Close resources
	        sourceWorkbook.close();
	        destWorkbook.close();

	        return "Data successfully written from 'Privacy Policy' (Col C, Row 4-207) to 'Content' (Col C, Row 3-206)!";
	    } catch (IOException e) {
	    	throw new ResourceNotFoundException("Error processing Excel file: " + e.getMessage());
	    }
 }
 private void resetFile() {
     try {
         File file = new File(outputFilePath);
         System.out.println("File will delete "+outputFilePath);
         // Delete the file if it exists
         if (file.exists()) {
             if (file.delete()) {
                 System.out.println("File deleted successfully.");
             } else {
                 System.out.println("Failed to delete the file.");
                 return;
             }
         }

         // Create a new empty file
         if (file.createNewFile()) {
             System.out.println("New file created: " + outputFilePath);
         } else {
             System.out.println("Failed to create the file.");
         }
     } catch (IOException e) {
         System.err.println("Error handling file: " + e.getMessage());
     }
 }

}

	 


