package enhancement;

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

public class SPSalesAmountAnalysisCreator {
	
	static final String SOURCE_DIR = "src\\main\\resources\\source\\enhancement";
	
	static final String TEMPLATE_NAME = "SP Sales Amount Analysis 20200715.xlsx";
	
	static final String DESTINATION_DIR = "src\\main\\resources\\result\\enhancement";
	
	static final String DESTINATION_FILE_NAME_FORMAT = "SP Sales Amount Analysis %s.xlsx";
	
	static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd HHmmssZ");
	
	static final String ODM_SP_SALES_AMOUNT_SHEET_NAME = "ODM SP Sales Amount"; 
	
	static final int ODM_SP_SALES_AMOUNT_MAX_ROW = 8;
	
	static final int ODM_SP_SALES_AMOUNT_START_ROW_INDEX = 1;
	
	static final int ODM_SP_SALES_AMOUNT_START_COLUMN_INDEX = 1;
	
	FileInputStream fileIS;
	
	FileOutputStream fileOS;
	
	Workbook workbook = null;
	
	DataFormat dataFormat = null;
	
	OdmSPSalesAmountCreator odmSPSalesAmountCreator = new OdmSPSalesAmountCreator();
	
	public static void main(String[] args) {
		try {
			SPSalesAmountAnalysisCreator creator = new SPSalesAmountAnalysisCreator();
			creator.generateWorkbook();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void generateWorkbook() throws IOException {
		try {
			if (!this.init()) {
				return;
			}
			this.initOdmSPSalesAmountCreator();
			odmSPSalesAmountCreator.generateWorkbook();
			this.outputWookbook();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			this.close();
		}
	}
	
	private boolean init() throws IOException {
    	Workbook workbook = this.getWorkbook();
    	if(workbook == null) {
    		return false;
    	}
    	this.setDataFormat();
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
	
	private void setDataFormat() {
		this.dataFormat = this.workbook.createDataFormat();
	}
	
	private void initOdmSPSalesAmountCreator() {
		Sheet odmSPSalesAmount = workbook.getSheet(ODM_SP_SALES_AMOUNT_SHEET_NAME);
		this.odmSPSalesAmountCreator.setTargetSheet(odmSPSalesAmount);
		this.odmSPSalesAmountCreator.setDataFormat(this.dataFormat);
		this.odmSPSalesAmountCreator.init();
	}
	
    private void outputWookbook() {
    	try { 
    		File excelFile = new File(DESTINATION_DIR,String.format(DESTINATION_FILE_NAME_FORMAT, simpleDateFormat.format(new Date())));
            this.fileOS = new FileOutputStream(excelFile); 
            this.workbook.write(this.fileOS); 
            System.out.println(String.format("%s written successfully on disk.", excelFile.getName())); 
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }
	
    private void close() throws IOException {
    	if(this.fileIS!=null) {
    		this.fileIS.close();
    	}
    	if(this.fileOS!=null) {
    		this.fileOS.close();
    	}
    }
}
