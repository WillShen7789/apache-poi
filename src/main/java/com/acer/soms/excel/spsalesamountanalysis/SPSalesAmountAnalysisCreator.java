package com.acer.soms.excel.spsalesamountanalysis;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.acer.soms.excel.odmspsalesamount.OdmSPSalesAmountCreator;
import com.acer.soms.excel.spsalesamountbar.SPSalesAmountBarChartCreator;

public class SPSalesAmountAnalysisCreator {
	
	static final String SOURCE_DIR = "src\\main\\resources\\source\\enhancement";
	
	static final String TEMPLATE_NAME = "SP Sales Amount Analysis 20200727.xlsx";
	
	static final String DESTINATION_DIR = "src\\main\\resources\\result\\enhancement";
	
	static final String DESTINATION_FILE_NAME_FORMAT = "SP Sales Amount Analysis %s.xlsx";
	
	static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd HHmmssZ");
	
	FileInputStream fileIS;
	
	FileOutputStream fileOS;
	
	Workbook workbook = null;
	
	OdmSPSalesAmountCreator odmSPSalesAmountCreator = new OdmSPSalesAmountCreator();
	
	SPSalesAmountBarChartCreator spSalesAmountBarChartCreator = new SPSalesAmountBarChartCreator();
	
	public static void main(String[] args) {
		SPSalesAmountAnalysisCreator creator = new SPSalesAmountAnalysisCreator();
		try {
			if (!creator.init()) {
				return;
			}
			creator.generateWorkbook();
			creator.outputWookbook();
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}
	
	private boolean init() throws IOException {
		try {
	    	Workbook workbook = this.getWorkbook();
	    	if(workbook == null) {
	    		return false;
	    	}
	    	odmSPSalesAmountCreator.init(workbook);
	    	spSalesAmountBarChartCreator.init(workbook);
		}catch(Exception e) {
			e.printStackTrace();
			return false;
		}
    	return true;
	}
	
	private Workbook getWorkbook() throws IOException {
		File templateFile = new File(SOURCE_DIR, TEMPLATE_NAME);
		if (!templateFile.canExecute()) {
			return null;
		}
		this.fileIS = new FileInputStream(templateFile);
		this.workbook = new XSSFWorkbook(fileIS);
		return this.workbook;
	}
	
	private void generateWorkbook() throws IOException {
		try {
			odmSPSalesAmountCreator.generateWorkbook();
			spSalesAmountBarChartCreator.generateWorkbook();
		} catch (Exception e) {
			e.printStackTrace();
			this.close();
		}
	}
	
    private void outputWookbook() throws IOException {
    	try { 
    		File excelFile = new File(DESTINATION_DIR,String.format(DESTINATION_FILE_NAME_FORMAT, simpleDateFormat.format(new Date())));
    		if(!excelFile.getParentFile().isDirectory()) {
    			if(!excelFile.getParentFile().mkdirs()) {
    				System.out.println(String.format("%s mkdirs failed.", excelFile.getParentFile().getAbsolutePath())); 
    				return;
    			}
    		}
    		this.fileOS = new FileOutputStream(excelFile); 
            this.workbook.write(this.fileOS); 
            System.out.println(String.format("%s written successfully on disk.", excelFile.getName())); 
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } finally {
        	this.close();
        }
    }
	
    private void close() throws IOException {
    	if(this.fileIS!=null) {
    		this.fileIS.close();
    		this.fileIS = null;
    	}
    	if(this.fileOS!=null) {
    		this.fileOS.close();
    		this.fileOS = null;
    	}
    }
}
