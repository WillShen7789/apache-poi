package exceltool;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	static final String TARGET_DIR = "src\\main\\resources\\source";
	
	static Map<String,Integer> columnNameToMaxLength = new LinkedHashMap<String,Integer>();
	
	static Map<Integer,String> tmpColumnNameList = new LinkedHashMap<Integer,String>();
	
	public static void main(String[] args) 
    { 
        try {
        	File targetDir = new File(TARGET_DIR);
        	if(targetDir.isDirectory() == false) {
        		return;
        	}
        	for(File targetFile : targetDir.listFiles()) {
        		handleFile(targetFile);
        	}
        	System.out.print("Total Count:" + columnNameToMaxLength.size() + "\t");
        	for(String columnName : columnNameToMaxLength.keySet()) {
        		System.out.print(String.format("%s(%d)", 
        			columnName, columnNameToMaxLength.get(columnName)) + "\t");
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
      
//      Map<String,Integer> columnNameToLength = new HashMap<String,Integer>();
      
      tmpColumnNameList.clear();
      
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
    	  if(columnNameToMaxLength.containsKey(columnName) == false) {
    		  columnNameToMaxLength.put(columnName, 20);
    		  tmpColumnNameList.put(m,columnName);
    	  }
//    	  columnNameToLength.put(columnName, 20);
    	  
      }

      handleMaxLength(sheet);

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
	
	private static void handleMaxLength(XSSFSheet sheet) {
	      for(int i=1;i<sheet.getLastRowNum();i++) {
	    	  Row targetRow = sheet.getRow(i);
	          for(int m=0;m<targetRow.getLastCellNum();m++) {
	        	  if(tmpColumnNameList.containsKey(m) == false) {
	        		  continue;
	        	  }
	        	  Cell cell = targetRow.getCell(m);
	        	  if(cell==null) {
	        		  continue;
	        	  }
	              switch (cell.getCellType()) { 
	              case Cell.CELL_TYPE_NUMERIC: 
//	                  System.out.print(cell.getNumericCellValue() + "\t"); 
	                  break; 
	              case Cell.CELL_TYPE_STRING:
	            	  String content = cell.getStringCellValue();
	            	  int len1 = content.length();
	            	  String targetName =  tmpColumnNameList.get(m);
	            	  if(columnNameToMaxLength.containsKey(targetName)) {
	            		 int len2 =  columnNameToMaxLength.get(tmpColumnNameList.get(m)).intValue();
	            		 if( len1 > len2) {
	            			 columnNameToMaxLength.put(tmpColumnNameList.get(m), len1);
	            		 }
	            	  }
	                  break; 
	              }
	          }
	      }
	}
	
}
