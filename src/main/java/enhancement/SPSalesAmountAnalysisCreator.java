package enhancement;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SPSalesAmountAnalysisCreator {
	
	static final String TARGET_DIR = "src\\main\\resources\\source\\enhancement";
	
	static final String TEMPLATE_NAME = "SP Sales Amount Analysis.xlsx";
	
	static final String DESTINATION_FILE_NAME_FORMAT = "SP Sales Amount Analysis %s.xlsx";
	
	static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyMMddHHmmssZ");
	
	public static void main(String[] args) 
    { 
        try {
        	File templateFile = new File(TARGET_DIR,TEMPLATE_NAME);
        	if(!templateFile.canExecute()) {
        		return;
        	}
        	
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }
	
	
	
	private static void handleFile(File file) throws IOException {
      FileInputStream fileIS = new FileInputStream(file); 
      XSSFWorkbook workbook = new XSSFWorkbook(fileIS); 
      XSSFSheet sheet = workbook.getSheet("Specifications"); 
      Row headColumnRow = sheet.getRow(0);
      System.out.print(String.format("Name:%s, Count:%d", file.getName(),headColumnRow.getLastCellNum()) + "\t");
      for(int m=0;m<headColumnRow.getLastCellNum();m++) {
    	  Cell cell = headColumnRow.getCell(m);
    	  String columnName = cell.getStringCellValue();
    	  columnName = substringEndAtFirstParentheses(columnName);
    	  columnName = convertFirstNumber(columnName);
    	  columnName = convertWhiteChar(columnName);
    	  columnName = columnName.toUpperCase();
    	  System.out.print(columnName + "\t");
      }

      System.out.println("");      
      workbook.close();
      fileIS.close(); 
	}
	
	private static String substringEndAtFirstParentheses(String columnName) {
		int endIdx = columnName.trim().indexOf("("); 
		if(endIdx >= 0) {
			return columnName.substring(0,endIdx).trim();
		}
		return columnName.trim();
	}
	
	private static String convertFirstNumber(String columnName) {
		String[] stringNumbers = new String[] {
		    "0","1","2","3","4","5","6","7","8","9"
		};
		String[] numberNames = new String[] {
			    "ZERO","ONE","TWO","THREE","FOUR",
			    "FIVE","SIX","SEVEN","EIGHT","NINE"
			};
	    for(int i=0;i<stringNumbers.length;i++) {
	    	if(columnName.indexOf(stringNumbers[i]) == 0) {
	    		columnName = columnName.replaceFirst("\\d",numberNames[i]+"_");
	    	}
	    }
		return columnName;
	}
	
	private static String convertWhiteChar(String columnName) {
		return columnName.replace(" ", "_").replace(".", "_").replace("/", "_");
	}
	
}
