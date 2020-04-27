package exceltool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSpliter {
	
	static final String TARGET_FILE = "src\\main\\resources\\source\\survey NOTEBOOK_WEEKLY_20170430.xlsx";
	
	static final String SOURCE_DIR = "src\\main\\resources\\source";
	
	static final String RESULT_DIR = "src\\main\\resources\\result";
	
	static final String FILE_EXTENSION = ".xlsx";
	
	static final int TARGET_FILE_LENGTH = 9700000;
	
	static final String TARGET_SHEET = "Specifications";
	
	public static void main(String[] args) {
		try {
			File targetDir = new File(SOURCE_DIR);
			if (targetDir.isDirectory() == false) {
				return;
			}
			for (File targetFile : targetDir.listFiles()) {
				if (targetFile.length() > TARGET_FILE_LENGTH) {
					// handleFile(targetFile);
				}
			}
			File targetFile = new File(TARGET_FILE);
			handleFile(targetFile);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static String readableFileSize(long size) {
        if(size <= 0) return "0";
        final String[] units = new String[] { "B", "kB", "MB", "GB", "TB" };
        int digitGroups = (int) (Math.log10(size)/Math.log10(1024));
        return new DecimalFormat("#,##0.#").format(size/Math.pow(1024, digitGroups)) +
        		" " + units[digitGroups];
    }
	
	private static void handleFile(File file) throws IOException {
		  
		  System.out.println(String.format("HandleFileName:%s, FileSize:%s", 
				  file.getName(),readableFileSize(file.length())));
		  
		  FileInputStream fileIn = new FileInputStream(file); 
	      XSSFWorkbook workBookIn = new XSSFWorkbook(fileIn); 
	      XSSFSheet sheetIn = workBookIn.getSheet(TARGET_SHEET); 
	      int helfRowNum = sheetIn.getLastRowNum() / 2;

	      XSSFWorkbook workbookOS1 = new XSSFWorkbook();
	      XSSFSheet sheetOS1 = workbookOS1.createSheet(TARGET_SHEET);
	      
	      List<String> columnNames = new ArrayList<String>();
	      Row row = sheetIn.getRow(0);
    	  for(int colIdx=0;colIdx<row.getLastCellNum();colIdx++) {
        	  Cell cell = row.getCell(colIdx);
        	  if(cell==null) {
        		 Object value =  getCellValue(cell);
        	  }
    	  }
    	  
	      for(int rowIdx=1;rowIdx<helfRowNum;rowIdx++) {
	    	  Row rowIn = sheetIn.getRow(rowIdx);
	    	  Row rowOut = sheetIn.createRow(rowIdx); 
	    	  for(int m=0;m<rowIn.getLastCellNum();m++) {
	        	  Cell cell = rowIn.getCell(m);
	        	  if(cell==null) {
	        		 Object value =  getCellValue(cell);
	        	  }
	    	  }
	      }
	      String fileName1 = getFileName(file.getName(),1);
		  System.out.println(fileName1);
	      
		  XSSFWorkbook workbookOS2 = new XSSFWorkbook();
	      for(int i=helfRowNum;i<sheetIn.getLastRowNum();i++) {
	      }
	      String fileName2 = getFileName(file.getName(),2);
	      System.out.println(fileName2);
	      
	      workBookIn.close();
	      fileIn.close();
	}
	
	
	private static String getFileName(String name, int number) {
		String fileName = name.substring(0,name.indexOf(FILE_EXTENSION));
		return String.format(fileName+"_%d", number) + FILE_EXTENSION;
	}
	
	private static Object getCellValue(Cell cell) {
		Object result = null;
        switch (cell.getCellType()) { 
	        case Cell.CELL_TYPE_NUMERIC: 
	        	result = cell.getNumericCellValue(); 
	            break; 
	        case Cell.CELL_TYPE_STRING: 
	        	result =  cell.getStringCellValue(); 
	            break; 
        }
        return result;
	}
	
	private static void writeExcelFile(XSSFWorkbook workbook, String name) 
			throws IOException {
		File result = new File(RESULT_DIR, name);
		FileOutputStream out = new FileOutputStream(result);
		workbook.write(out);
		out.close();
		System.out.println(String.format("%s written successfully on disk.", result.getName()));
	}
	
	public static void test(String[] args) 
    { 
        // Blank workbook 
        XSSFWorkbook workbook = new XSSFWorkbook(); 
  
        // Create a blank sheet 
        XSSFSheet sheet = workbook.createSheet("student Details"); 
  
        // This data needs to be written (Object[]) 
        Map<String, Object[]> data = new TreeMap<String, Object[]>(); 
        data.put("1", new Object[]{ "ID", "NAME", "LASTNAME" }); 
        data.put("2", new Object[]{ 1, "Pankaj", "Kumar" }); 
        data.put("3", new Object[]{ 2, "Prakashni", "Yadav" }); 
        data.put("4", new Object[]{ 3, "Ayan", "Mondal" }); 
        data.put("5", new Object[]{ 4, "Virat", "kohli" }); 
  
        // Iterate over data and write to sheet 
        Set<String> keyset = data.keySet(); 
        int rownum = 0; 
        for (String key : keyset) { 
            // this creates a new row in the sheet 
            Row row = sheet.createRow(rownum++); 
            Object[] objArr = data.get(key); 
            int cellnum = 0; 
            for (Object obj : objArr) { 
                // this line creates a cell in the next column of that row 
                Cell cell = row.createCell(cellnum++); 
                if (obj instanceof String) 
                    cell.setCellValue((String)obj); 
                else if (obj instanceof Integer) 
                    cell.setCellValue((Integer)obj); 
            } 
        } 
        try { 
            // this Writes the workbook gfgcontribute 
            FileOutputStream out = new FileOutputStream(new File("gfgcontribute.xlsx")); 
            workbook.write(out); 
            out.close(); 
            System.out.println("gfgcontribute.xlsx written successfully on disk."); 
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }
}
