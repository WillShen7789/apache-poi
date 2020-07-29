package com.acer.soms.excel.spsalesamountbar;

import java.math.BigDecimal;
import java.util.List;

public class SPSalesAmountBarCharTVos {
		
	private List<String> shipDateMonth;
	private List<BigDecimal> bsAmount;
	private List<BigDecimal> odmPartsAmount;
	private List<BigDecimal> fgAmount;
	
	private int rowSize = 0;

	public List<String> getShipDateMonth() {
		return shipDateMonth;
	}

	public void setShipDateMonth(List<String> shipDateMonth) {
		this.shipDateMonth = shipDateMonth;
	}

	public List<BigDecimal> getBsAmount() {
		return bsAmount;
	}

	public void setBsAmount(List<BigDecimal> bsAmount) {
		this.bsAmount = bsAmount;
	}

	public List<BigDecimal> getOdmPartsAmount() {
		return odmPartsAmount;
	}

	public void setOdmPartsAmount(List<BigDecimal> odmPartsAmount) {
		this.odmPartsAmount = odmPartsAmount;
	}

	public List<BigDecimal> getFgAmount() {
		return fgAmount;
	}

	public void setFgAmount(List<BigDecimal> fgAmount) {
		this.fgAmount = fgAmount;
	}

	public int getRowSize() {
		return rowSize;
	}

	public void setRowSize(int rowSize) {
		this.rowSize = rowSize;
	}
	
}
