package enhancement;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

public class OdmSPSalesAmountCreator {

	static final String ODM_SP_SALES_AMOUNT_SHEET_NAME = "ODM SP Sales Amount"; 
	
	static final String MODEL_STYLE_KEY = "modelStyle";
	static final String AMOUNT_STYLE_KEY = "amountStyle";
	static final String PERCENT_STYLE_KEY = "percentStyle";
	
	static final int START_ROW_INDEX = 1;
	static final int START_COLUMN_INDEX = 1;
	static final int AMOUNT_START_ROW_INDEX = 2;
	
	static final String SP_SALES_AMOUNT_ANALYSIS_SHEET_NAME = "SP Sales Amount Analysis";
	static final int SP_SALES_AMOUNT_ANALYSIS_AMOUNT_START_ROW_NUM = 14;
	static final int SP_SALES_AMOUNT_ANALYSIS_AMOUNT_END_ROW_NUM = 19;
	static final int SP_SALES_AMOUNT_ANALYSIS_PERCENT_ROW_NUM = 28;
	static final int SP_SALES_AMOUNT_ANALYSIS_AMOUNT_ROW_OFFSET = SP_SALES_AMOUNT_ANALYSIS_AMOUNT_START_ROW_NUM
			- AMOUNT_START_ROW_INDEX;
	
	private DataFormat dataFormat = null;
	
	private Sheet targetSheet = null;

	private Map<String,CellStyle> styleMap = new HashMap<String,CellStyle>();
	
	public void init(Workbook workbook) {
		
		targetSheet = workbook.getSheet(ODM_SP_SALES_AMOUNT_SHEET_NAME);
		this.dataFormat = workbook.createDataFormat();
		
		Row row = this.targetSheet.getRow(1);
		Cell cell = row.getCell(1);
		CellStyle style = cell.getCellStyle();
		styleMap.put(MODEL_STYLE_KEY, style);
		
		row = this.targetSheet.getRow(2);
		cell = row.getCell(1);
		style = cell.getCellStyle();
		styleMap.put(AMOUNT_STYLE_KEY, style);
		
		row = this.targetSheet.getRow(8);
		cell = row.getCell(1);
		style = cell.getCellStyle();
		styleMap.put(PERCENT_STYLE_KEY, style);
	}
	
	public void generateWorkbook() {
//    	List<OdmSPSalesAmountVo> datas = this.getData();
//    	OdmSPSalesAmountTVos tVos = this.transferData(datas);
		List<String> shipDateMonths = this.getData();
    	this.setDataToSheet(shipDateMonths);
	}
	
//	private List<OdmSPSalesAmountVo> getData(){
//		
//		List<OdmSPSalesAmountVo> vos = new ArrayList<OdmSPSalesAmountVo>();
//		
//		OdmSPSalesAmountVo vo = new OdmSPSalesAmountVo();
//		vo.setShipDateMonth("2020-01");
//		vo.setNbAmount(new BigDecimal("1.47"));
//		vo.setDtAmount(new BigDecimal("0.18"));
//		vo.setSvrAmount(new BigDecimal("0.00"));
//		vo.setMntAmount(new BigDecimal("0.12"));
//		vo.setPrjAmount(new BigDecimal("0.03"));
//		vo.setTotalAmount(new BigDecimal("1.86"));
//		vo.setPercent(new BigDecimal("0.00439732095238833"));
//		vos.add(vo);
//		
//		vo = new OdmSPSalesAmountVo();
//		vo.setShipDateMonth("2020-02");
//		vo.setNbAmount(new BigDecimal("1.47"));
//		vo.setDtAmount(new BigDecimal("0.18"));
//		vo.setSvrAmount(new BigDecimal("0.00"));
//		vo.setMntAmount(new BigDecimal("0.12"));
//		vo.setPrjAmount(new BigDecimal("0.03"));
//		vo.setTotalAmount(new BigDecimal("1.86"));
//		vo.setPercent(new BigDecimal("0.00439732095238833"));
//		vos.add(vo);
//		
//		vo = new OdmSPSalesAmountVo();
//		vo.setShipDateMonth("2020-03");
//		vo.setNbAmount(new BigDecimal("1.47"));
//		vo.setDtAmount(new BigDecimal("0.18"));
//		vo.setSvrAmount(new BigDecimal("0.00"));
//		vo.setMntAmount(new BigDecimal("0.12"));
//		vo.setPrjAmount(new BigDecimal("0.03"));
//		vo.setTotalAmount(new BigDecimal("1.86"));
//		vo.setPercent(new BigDecimal("0.00439732095238833"));
//		vos.add(vo);
//		
//		vo = new OdmSPSalesAmountVo();
//		vo.setShipDateMonth("2020-04");
//		vo.setNbAmount(new BigDecimal("1.47"));
//		vo.setDtAmount(new BigDecimal("0.18"));
//		vo.setSvrAmount(new BigDecimal("0.00"));
//		vo.setMntAmount(new BigDecimal("0.12"));
//		vo.setPrjAmount(new BigDecimal("0.03"));
//		vo.setTotalAmount(new BigDecimal("1.86"));
//		vo.setPercent(new BigDecimal("0.00439732095238833"));
//		vos.add(vo);
//		
//		return vos;
//	}
	
//	private OdmSPSalesAmountTVos transferData(List<OdmSPSalesAmountVo> data) {
//		OdmSPSalesAmountTVos tvos = new OdmSPSalesAmountTVos();
//		if(data==null || data.size() <= 0) {
//			return tvos;
//		}
//		for(int i=0;i<data.size();i++) {
//			OdmSPSalesAmountVo vo = data.get(i);;
//			if(tvos.getShipDateMonth() == null) {
//				List<String> shipDateMonths = new ArrayList<String>();
//				shipDateMonths.add(vo.getShipDateMonth());
//				tvos.setShipDateMonth(shipDateMonths);
//			} else {
//				tvos.getShipDateMonth().add(vo.getShipDateMonth());
//			}
//			if(tvos.getNbAmount() == null) {
//				List<BigDecimal> nbAmounts = new ArrayList<BigDecimal>();
//				nbAmounts.add(vo.getNbAmount());
//				tvos.setNbAmount(nbAmounts);
//			} else {
//				tvos.getNbAmount().add(vo.getNbAmount());
//			}
//			if(tvos.getDtAmount() == null) {
//				List<BigDecimal> dtAmounts = new ArrayList<BigDecimal>();
//				dtAmounts.add(vo.getDtAmount());
//				tvos.setDtAmount(dtAmounts);
//			} else {
//				tvos.getDtAmount().add(vo.getDtAmount());
//			}
//			if(tvos.getSvrAmount() == null) {
//				List<BigDecimal> svrAmounts = new ArrayList<BigDecimal>();
//				svrAmounts.add(vo.getSvrAmount());
//				tvos.setSvrAmount(svrAmounts);
//			} else {
//				tvos.getSvrAmount().add(vo.getSvrAmount());
//			}
//			if(tvos.getMntAmount() == null) {
//				List<BigDecimal> mntAmounts = new ArrayList<BigDecimal>();
//				mntAmounts.add(vo.getMntAmount());
//				tvos.setMntAmount(mntAmounts);
//			} else {
//				tvos.getMntAmount().add(vo.getMntAmount());
//			}
//			if(tvos.getPrjAmount() == null) {
//				List<BigDecimal> prjAmounts = new ArrayList<BigDecimal>();
//				prjAmounts.add(vo.getPrjAmount());
//				tvos.setPrjAmount(prjAmounts);
//			} else {
//				tvos.getPrjAmount().add(vo.getPrjAmount());
//			}
//			if(tvos.getTotalAmount() == null) {
//				List<BigDecimal> totalAmounts = new ArrayList<BigDecimal>();
//				totalAmounts.add(vo.getTotalAmount());
//				tvos.setTotalAmount(totalAmounts);
//			} else {
//				tvos.getTotalAmount().add(vo.getTotalAmount());
//			}
//			if(tvos.getPercent() == null) {
//				List<BigDecimal> percents = new ArrayList<BigDecimal>();
//				percents.add(vo.getPercent());
//				tvos.setPercent(percents);
//			} else {
//				tvos.getPercent().add(vo.getPercent());
//			}			
//		}
//		return tvos;
//	}
	
	private List<String> getData(){
		List<String> shipDateMonths = new ArrayList<String>();
		shipDateMonths.add("2020-01");
		shipDateMonths.add("2020-02");
		shipDateMonths.add("2020-03");
		shipDateMonths.add("2020-04");
	    return shipDateMonths;
	}
	
	private void setDataToSheet(List<String> shipDateMonths) {
    	
    	int rowIdx = START_ROW_INDEX;
		Row row = this.targetSheet.getRow(rowIdx);
		int colIdx = START_COLUMN_INDEX;
		
//		List<String> shipDateMonths = tVos.getShipDateMonth();
		this.setShipDateMonths(shipDateMonths,row,colIdx);
		rowIdx++;
    	
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> nbAmounts = tVos.getNbAmount();
		this.setAmounts(shipDateMonths, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> dtAmounts = tVos.getDtAmount();
		this.setAmounts(shipDateMonths, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> svrAmounts = tVos.getSvrAmount();
		this.setAmounts(shipDateMonths, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> mntAmounts = tVos.getMntAmount();
		this.setAmounts(shipDateMonths, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> prjAmounts = tVos.getPrjAmount();
		this.setAmounts(shipDateMonths, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> totalAmounts = tVos.getTotalAmount();
		this.setAmounts(shipDateMonths, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
//		List<BigDecimal> percents = tVos.getPercent();
		this.setPercents(shipDateMonths, row, colIdx);
		rowIdx++;
	}
	
	private void setShipDateMonths(List<String> shipDateMonths,Row row, int colIdx) {
		for (int m = 0; m < shipDateMonths.size(); m++) {
			if(m==0) {
				Cell cell = row.getCell(colIdx);
				cell.setCellValue(shipDateMonths.get(m));
			} else {
				Cell cell = row.getCell(colIdx);
				if(cell==null) {
					cell = row.createCell(colIdx);
				}
				cell.setCellStyle(this.styleMap.get(MODEL_STYLE_KEY));
				cell.setCellValue(shipDateMonths.get(m));
			}
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellStyle(this.styleMap.get(MODEL_STYLE_KEY));
		lastColCell.setCellValue("Total");
	}
	
	private void setAmounts(List<String> shipDateMonths, Row row, int colIdx) {
		int shipDateMonthsSize = shipDateMonths.size();
		for (int m = 0; m < shipDateMonthsSize; m++) {
			Cell cell = row.getCell(colIdx);
			if(cell==null) {
				cell = row.createCell(colIdx);
			}
			cell.setCellType(CellType.FORMULA);
			cell.setCellStyle(this.styleMap.get(AMOUNT_STYLE_KEY));
			String colString = CellReference.convertNumToColString(colIdx);
			int refRowNum = row.getRowNum() + SP_SALES_AMOUNT_ANALYSIS_AMOUNT_ROW_OFFSET;
			String refCell = String.format("'%s'!%s%d", SP_SALES_AMOUNT_ANALYSIS_SHEET_NAME,colString,refRowNum);
			cell.setCellFormula(String.format("IFERROR(%s,0)",refCell));
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellType(CellType.FORMULA);
		lastColCell.setCellStyle(this.styleMap.get(AMOUNT_STYLE_KEY));
		CellReference cellRef1 = new CellReference(row.getRowNum(),START_COLUMN_INDEX);
		CellReference cellRef2 = new CellReference(row.getRowNum(),START_COLUMN_INDEX + shipDateMonthsSize -1);
		lastColCell.setCellFormula(String.format("SUM(%s:%s)", cellRef1.formatAsString(),cellRef2.formatAsString()));
	}
	
	private void setPercents(List<String> shipDateMonths, Row row, int colIdx) {
		int shipDateMonthsSize = shipDateMonths.size();
		for (int m = 0; m < shipDateMonthsSize; m++) {
		    Cell cell = row.getCell(colIdx);
			if(cell==null) {
				cell = row.createCell(colIdx);
			}
			cell.setCellType(CellType.FORMULA);
		    CellStyle style = this.styleMap.get(PERCENT_STYLE_KEY);
			style.setDataFormat(this.dataFormat.getFormat("0.00%"));
			cell.setCellStyle(style);
			String colString = CellReference.convertNumToColString(colIdx);
			int refRowNum = SP_SALES_AMOUNT_ANALYSIS_PERCENT_ROW_NUM;
			String refCell = String.format("'%s'!%s%d", SP_SALES_AMOUNT_ANALYSIS_SHEET_NAME,colString,refRowNum);
			cell.setCellFormula(String.format("IFERROR(%s,0)",refCell));
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellType(CellType.FORMULA);
	    CellStyle style = this.styleMap.get(PERCENT_STYLE_KEY);
		style.setDataFormat(this.dataFormat.getFormat("0.00%"));
		lastColCell.setCellStyle(style);
		CellReference cellRef1 = new CellReference(row.getRowNum(),START_COLUMN_INDEX);
		CellReference cellRef2 = new CellReference(row.getRowNum(),START_COLUMN_INDEX + shipDateMonthsSize -1);
		lastColCell.setCellFormula(String.format("SUM(%s:%s)", cellRef1.formatAsString(),cellRef2.formatAsString()));
	}
}
