package enhancement;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
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
        	System.out.println("Template File Name : " + templateFile.getName());
        	
        	SPSalesAmountAnalysisCreator creator = new SPSalesAmountAnalysisCreator();
        	
        	XSSFWorkbook workbook = creator.getWorkbook(templateFile);
        	
        	XSSFSheet odmSPSalesAmount = workbook.getSheet("ODM SP Sales Amount");
        	
        	
        	List<OdmSPSalesAmountVo> datas = creator.getData();
        	
        	int startCol = 1;
        	for(int i=0;i<datas.size();i++) {
        		OdmSPSalesAmountVo data = datas.get(i);
        		XSSFRow row = odmSPSalesAmount.getRow(1);
        		XSSFCell cell =  row.getCell(startCol);
        	}
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }
	
	private XSSFWorkbook getWorkbook(File templateFile) throws IOException {
	      FileInputStream fileIS = new FileInputStream(templateFile); 
	      return new XSSFWorkbook(fileIS); 
	}
	
	private List<OdmSPSalesAmountVo> getData(){
		List<OdmSPSalesAmountVo> vos = new ArrayList<OdmSPSalesAmountVo>();
		
		OdmSPSalesAmountVo vo = new OdmSPSalesAmountVo();
		vo.setShipDateMonth("2020-01");
		vo.setNbAmount(new BigDecimal("1.47"));
		vo.setDtAmount(new BigDecimal("0.18"));
		vo.setSvrAmount(new BigDecimal("0.00"));
		vo.setMntAmount(new BigDecimal("0.12"));
		vo.setPrjAmount(new BigDecimal("0.03"));
		vo.setTotalAmount(new BigDecimal("1.86"));
		vo.setPercent(new BigDecimal("0.44"));
		
		vos.add(vo);
		
		return vos;
	}
	
	private OdmSPSalesAmountTVos transferData(List<OdmSPSalesAmountVo> data) {
		OdmSPSalesAmountTVos tvos = new OdmSPSalesAmountTVos();
		if(data==null || data.size() <= 0) {
			return tvos;
		}
		for(int i=0;i<data.size();i++) {
			OdmSPSalesAmountVo vo = data.get(i);;
			if(tvos.getShipDateMonth() == null) {
				List<String> shipDateMonths = new ArrayList<String>();
				shipDateMonths.add(vo.getShipDateMonth());
				tvos.setShipDateMonth(shipDateMonths);
			} else {
				tvos.getShipDateMonth().add(vo.getShipDateMonth());
			}
			if(tvos.getNbAmount() == null) {
				List<BigDecimal> nbAmounts = new ArrayList<BigDecimal>();
				nbAmounts.add(vo.getNbAmount());
				tvos.setNbAmount(nbAmounts);
			} else {
//				tvos.getShipDateMonth().add(vo.get);
			}
		}
		return tvos;
	}


	
}
