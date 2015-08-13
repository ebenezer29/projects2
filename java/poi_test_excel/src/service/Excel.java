package service;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Excel {

	public static void createExcel (String fileName) throws Exception {
		
		File file = new File(fileName);
		
        if (!file.exists()){
	    	Workbook workbook = new XSSFWorkbook();
	    	Sheet sheet = workbook.createSheet("sheet1");
	    	
	    	FileOutputStream fileOut = new FileOutputStream(fileName);
	    	workbook.write(fileOut);
	    	fileOut.close();
        } 
	}
	
	public static void writeToExcel (int row, int column, String value, String fileName) throws Exception {
        
        FileInputStream fileIn = new FileInputStream(new File(fileName));
        XSSFWorkbook workbook = new XSSFWorkbook(fileIn);  
        XSSFSheet worksheet = workbook.getSheetAt(0);   
        Row myRow = null;
        Cell cell = null;
        
        int numOfRows = worksheet.getPhysicalNumberOfRows(); 
        
        if (row >= numOfRows || numOfRows == 0){
        	myRow = worksheet.createRow(row);
        } else {
        	myRow = worksheet.getRow(row);
        }
        
        Cell myCell = myRow.createCell(column);   
        cell = worksheet.getRow(row).getCell(column);
        cell.setCellValue(value);  
        fileIn.close();
        FileOutputStream fileOut =new FileOutputStream(new File(fileName));
        workbook.write(fileOut);    
        fileOut.close();   
	}
	
	public static String readFromExcel (int row, int column, String fileName) throws Exception {
		
		String cellValue = null;
		File file = new File(fileName);
		
        if (!file.exists()){
        	System.out.println("the file requested doesn't exist");
        	cellValue = "null";
        } else {
			FileInputStream fileIn = new FileInputStream(new File(fileName));
			XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
			XSSFSheet worksheet = workbook.getSheetAt(0);
			
			Row myRow = null;
			
	        int numOfRows = worksheet.getPhysicalNumberOfRows(); 
	        
	        if (row >= numOfRows || numOfRows == 0){
	        	System.out.println("the row requested doesn't exist");
	        	cellValue = "null";
	        } else {
	        	myRow = worksheet.getRow(row);
	    		Cell cell = null; 
	    		cell = myRow.getCell(column);
	    		cellValue = cell.getStringCellValue();
	        }
			
			fileIn.close();
			workbook.close();
        }
        
        return cellValue;
	}	
}