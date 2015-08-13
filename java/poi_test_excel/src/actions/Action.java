package actions;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import service.Excel; 

public class Action {

	private static WebElement element = null;
	
	public static WebElement runTest(WebDriver driver) throws Exception {

		String inputfileName = "workbook.xlsx";
		String outputfileName = "workbook.xlsx";
		String value = null;
		
		Excel.createExcel(outputfileName); // if the file exists this will be skipped
		
		Excel.writeToExcel(0, 0, "write this", outputfileName);
		Excel.writeToExcel(0, 2, "write this", outputfileName);
		Excel.writeToExcel(1, 3, "write this", outputfileName);
	    Excel.writeToExcel(1, 1, "write this", outputfileName);
	    Excel.writeToExcel(2, 2, "write this", outputfileName);
	    
	    value = Excel.readFromExcel(2, 2, inputfileName); /* if you attempt to read from a file that doesn't exist 
	    													or a row that doesn't exist this will return null */
	    System.out.println(value);
	    
	    return element;
	}
}
