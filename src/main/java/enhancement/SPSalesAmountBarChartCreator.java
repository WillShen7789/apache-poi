package enhancement;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

public class SPSalesAmountBarChartCreator {

	static final String SP_SALES_AMOUNT_BAR_SHEET_NAME = "SP Sales Amount Bar chart"; 
	
	static final String MODEL_INTER_STYLE_KEY = "modelInterStyle";
	
	static final String MODEL_LAST_STYLE_KEY = "modelLastStyle";
	
	static final String AMOUNT_INTER_STYLE_KEY = "amountInterStyle";
	
	static final String AMOUNT_LAST_STYLE_KEY = "amountLastStyle";
	
	static final String FG_INTER_STYLE_KEY = "fgInterStyle";
	
	static final String FG_LAST_STYLE_KEY = "fgLastStyle";
	
	static final String PERCENT_INTER_STYLE_KEY = "percentInterStyle";
	
	static final String PERCENT_LAST_STYLE_KEY = "percentLastStyle";
	
	static final int TITLE_ROW_INDEX = 1;
	
	static final int TITLE_COLUMN_INDEX = 1;
	
	static final int START_ROW_INDEX = 2;
	
	static final int START_COLUMN_INDEX = 2;
	
	private DataFormat dataFormat = null;
	
	private Sheet targetSheet = null;

	private Map<String,CellStyle> styleMap = new HashMap<String,CellStyle>();
	
	public void init(Workbook workbook) {
		
		targetSheet = workbook.getSheet(SP_SALES_AMOUNT_BAR_SHEET_NAME);
		this.dataFormat = workbook.createDataFormat();
		
		Row row = this.targetSheet.getRow(2);
		Cell cell = row.getCell(2);
		CellStyle style = cell.getCellStyle();
		styleMap.put(MODEL_INTER_STYLE_KEY, style);
		
		cell = row.getCell(3);
		style = cell.getCellStyle();
		styleMap.put(MODEL_LAST_STYLE_KEY, style);
		
		row = this.targetSheet.getRow(3);
		cell = row.getCell(2);
		style = cell.getCellStyle();
		styleMap.put(AMOUNT_INTER_STYLE_KEY, style);
		
		cell = row.getCell(3);
		style = cell.getCellStyle();
		styleMap.put(AMOUNT_LAST_STYLE_KEY, style);
		
		row = this.targetSheet.getRow(6);
		cell = row.getCell(2);
		style = cell.getCellStyle();
		styleMap.put(FG_INTER_STYLE_KEY, style);
		
		cell = row.getCell(3);
		style = cell.getCellStyle();
		styleMap.put(FG_LAST_STYLE_KEY, style);
		
		row = this.targetSheet.getRow(7);
		cell = row.getCell(2);
		style = cell.getCellStyle();
		styleMap.put(PERCENT_INTER_STYLE_KEY, style);
		
		cell = row.getCell(3);
		style = cell.getCellStyle();
		styleMap.put(PERCENT_LAST_STYLE_KEY, style);
	}
	
	public void generateWorkbook() {
    	List<SPSalesAmountBarCharVo> datas = this.getData();
    	SPSalesAmountBarCharTVos tVos = this.transferData(datas);
    	this.setTitleStyle(tVos.getRowSize());
    	this.setDataToSheet(tVos);
	}
	
	private List<SPSalesAmountBarCharVo> getData(){
		
		List<SPSalesAmountBarCharVo> vos = new ArrayList<SPSalesAmountBarCharVo>();
		
		SPSalesAmountBarCharVo vo = new SPSalesAmountBarCharVo();
		vo.setShipDateMonth("2020-01");
		vo.setBsAmount(new BigDecimal("0.93426605"));
		vo.setOdmPartsAmount(new BigDecimal("1.86"));
		vo.setFgAmount(new BigDecimal("423.53"));
		vos.add(vo);
		
		vo = new SPSalesAmountBarCharVo();
		vo.setShipDateMonth("2020-02");
		vo.setBsAmount(new BigDecimal("1.05057919"));
		vo.setOdmPartsAmount(new BigDecimal("0.74"));
		vo.setFgAmount(new BigDecimal("423.53"));
		vos.add(vo);
		
		vo = new SPSalesAmountBarCharVo();
		vo.setShipDateMonth("2020-03");
		vo.setBsAmount(new BigDecimal("1.41310465"));
		vo.setOdmPartsAmount(new BigDecimal("3.43"));
		vo.setFgAmount(new BigDecimal("752.21"));
		vos.add(vo);
		
		vo = new SPSalesAmountBarCharVo();
		vo.setShipDateMonth("2020-04");
		vo.setBsAmount(new BigDecimal("0"));
		vo.setOdmPartsAmount(new BigDecimal("2.16"));
		vo.setFgAmount(new BigDecimal("0"));
		vos.add(vo);
		
		return vos;
	}
	
	private SPSalesAmountBarCharTVos transferData(List<SPSalesAmountBarCharVo> data) {
		SPSalesAmountBarCharTVos tvos = new SPSalesAmountBarCharTVos();
		if(data==null || data.size() <= 0) {
			return tvos;
		}
		tvos.setRowSize(data.size());
		for(int i=0;i<data.size();i++) {
			SPSalesAmountBarCharVo vo = data.get(i);;
			if(tvos.getShipDateMonth() == null) {
				List<String> shipDateMonths = new ArrayList<String>();
				shipDateMonths.add(vo.getShipDateMonth());
				tvos.setShipDateMonth(shipDateMonths);
			} else {
				tvos.getShipDateMonth().add(vo.getShipDateMonth());
			}
			if(tvos.getBsAmount() == null) {
				List<BigDecimal> bsAmount = new ArrayList<BigDecimal>();
				bsAmount.add(vo.getBsAmount());
				tvos.setBsAmount(bsAmount);
			} else {
				tvos.getBsAmount().add(vo.getBsAmount());
			}
			if(tvos.getOdmPartsAmount() == null) {
				List<BigDecimal> odmPartsAmount = new ArrayList<BigDecimal>();
				odmPartsAmount.add(vo.getOdmPartsAmount());
				tvos.setOdmPartsAmount(odmPartsAmount);
			} else {
				tvos.getOdmPartsAmount().add(vo.getOdmPartsAmount());
			}
			if(tvos.getFgAmount() == null) {
				List<BigDecimal> fgAmount = new ArrayList<BigDecimal>();
				fgAmount.add(vo.getFgAmount());
				tvos.setFgAmount(fgAmount);
			} else {
				tvos.getFgAmount().add(vo.getFgAmount());
			}
		}
		return tvos;
	}
	
	private void setTitleStyle(int dataSize) {
		int titleRowIdx = TITLE_ROW_INDEX;
		int titleColumnIdx = TITLE_COLUMN_INDEX;
		int lastColumnIndex = TITLE_COLUMN_INDEX + dataSize + 1;
		Row tRow = this.targetSheet.getRow(TITLE_ROW_INDEX);
		Cell tCell = tRow.getCell(TITLE_COLUMN_INDEX);
		CellStyle style = tCell.getCellStyle();
		for (int i = titleColumnIdx + 1; i <= lastColumnIndex; i++) {
			Cell tCell2 = tRow.getCell(i);
			if (tCell2 == null) {
				tCell2 = tRow.createCell(i);
			}
			tCell2.setCellStyle(style);
		}
		this.targetSheet
				.addMergedRegion(new CellRangeAddress(titleRowIdx, titleRowIdx, titleColumnIdx, lastColumnIndex));
	}
	
	private void setDataToSheet(SPSalesAmountBarCharTVos tVos) {
		
    	int rowIdx = START_ROW_INDEX;
		int colIdx = START_COLUMN_INDEX;
		
		Row row = this.targetSheet.getRow(rowIdx);
		
		List<String> shipDateMonths = tVos.getShipDateMonth();
		this.setShipDateMonths(shipDateMonths,row,colIdx);
		rowIdx++;
    	
		row = this.targetSheet.getRow(rowIdx);
		List<BigDecimal> nbAmounts = tVos.getBsAmount();
		this.setAmounts(nbAmounts, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
		List<BigDecimal> dtAmounts = tVos.getOdmPartsAmount();
		this.setAmounts(dtAmounts, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
		this.setAlls(tVos.getRowSize(), row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
		List<BigDecimal> mntAmounts = tVos.getFgAmount();
		this.setFGs(mntAmounts, row, colIdx);
		rowIdx++;
		
		row = this.targetSheet.getRow(rowIdx);
		this.setPercents(tVos.getRowSize(), row, colIdx);
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
				cell.setCellStyle(this.styleMap.get(MODEL_INTER_STYLE_KEY));
				cell.setCellValue(shipDateMonths.get(m));
			}
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellStyle(this.styleMap.get(MODEL_LAST_STYLE_KEY));
		lastColCell.setCellValue("Total");
	}
	
	private void setAmounts(List<BigDecimal> amounts, Row row, int colIdx) {
		for (int m = 0; m < amounts.size(); m++) {
			Cell cell = row.getCell(colIdx);
			if(cell==null) {
				cell = row.createCell(colIdx);
			}
			cell.setCellStyle(this.styleMap.get(AMOUNT_INTER_STYLE_KEY));
			BigDecimal nbAmount = amounts.get(m);
			cell.setCellValue(nbAmount.doubleValue());
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellType(CellType.FORMULA);
		lastColCell.setCellStyle(this.styleMap.get(AMOUNT_LAST_STYLE_KEY));
		CellReference cellRef1 = new CellReference(row.getRowNum(),START_COLUMN_INDEX);
		CellReference cellRef2 = new CellReference(row.getRowNum(),START_COLUMN_INDEX + amounts.size() -1);
		lastColCell.setCellFormula(String.format("SUM(%s:%s)", cellRef1.formatAsString(),cellRef2.formatAsString()));
	}
	
	private void setAlls(int dataSize, Row row, int colIdx) {
		for (int m = 0; m < dataSize; m++) {
			Cell cell = row.getCell(colIdx);
			if(cell==null) {
				cell = row.createCell(colIdx);
			}
			cell.setCellType(CellType.FORMULA);
			cell.setCellStyle(this.styleMap.get(AMOUNT_INTER_STYLE_KEY));
			CellReference cRef1 = new CellReference(row.getRowNum()-2,colIdx);
			CellReference cRef2 = new CellReference(row.getRowNum()-1,colIdx);
			cell.setCellFormula(String.format("SUM(%s:%s)", cRef1.formatAsString(),cRef2.formatAsString()));
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellType(CellType.FORMULA);
		lastColCell.setCellStyle(this.styleMap.get(AMOUNT_LAST_STYLE_KEY));
		CellReference cellRef1 = new CellReference(row.getRowNum(),START_COLUMN_INDEX);
		CellReference cellRef2 = new CellReference(row.getRowNum(),START_COLUMN_INDEX + dataSize -1);
		lastColCell.setCellFormula(String.format("SUM(%s:%s)", cellRef1.formatAsString(),cellRef2.formatAsString()));
	}
	
	private void setFGs(List<BigDecimal> amounts, Row row, int colIdx) {
		for (int m = 0; m < amounts.size(); m++) {
			Cell cell = row.getCell(colIdx);
			if(cell==null) {
				cell = row.createCell(colIdx);
			}
			cell.setCellStyle(this.styleMap.get(FG_INTER_STYLE_KEY));
			BigDecimal nbAmount = amounts.get(m);
			cell.setCellValue(nbAmount.doubleValue());
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellType(CellType.FORMULA);
		lastColCell.setCellStyle(this.styleMap.get(FG_LAST_STYLE_KEY));
		CellReference cellRef1 = new CellReference(row.getRowNum(),START_COLUMN_INDEX);
		CellReference cellRef2 = new CellReference(row.getRowNum(),START_COLUMN_INDEX + amounts.size() -1);
		lastColCell.setCellFormula(String.format("SUM(%s:%s)", cellRef1.formatAsString(),cellRef2.formatAsString()));
	}
	
	private void setPercents(int dataSize, Row row, int colIdx) {
		for (int m = 0; m < dataSize; m++) {
			Cell cell = row.getCell(colIdx);
			if(cell==null) {
				cell = row.createCell(colIdx);
			}
			cell.setCellType(CellType.FORMULA);
			cell.setCellStyle(this.styleMap.get(PERCENT_INTER_STYLE_KEY));
			CellReference cRef1 = new CellReference(row.getRowNum()-2,colIdx);
			CellReference cRef2 = new CellReference(row.getRowNum()-1,colIdx);
			cell.setCellFormula(String.format("%s/%s", cRef1.formatAsString(),cRef2.formatAsString()));
			colIdx++;
		}
		Cell lastColCell = row.getCell(colIdx);
		if(lastColCell==null) {
			lastColCell = row.createCell(colIdx);
		}
		lastColCell.setCellType(CellType.FORMULA);
	    CellStyle style = this.styleMap.get(PERCENT_LAST_STYLE_KEY);
		style.setDataFormat(this.dataFormat.getFormat("0.00%"));
		lastColCell.setCellStyle(style);
		CellReference cellRef1 = new CellReference(row.getRowNum(),START_COLUMN_INDEX);
		CellReference cellRef2 = new CellReference(row.getRowNum(),START_COLUMN_INDEX + dataSize -1);
		lastColCell.setCellFormula(String.format("%s/%s", cellRef1.formatAsString(),cellRef2.formatAsString()));
	}
}
