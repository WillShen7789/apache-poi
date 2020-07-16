package enhancement;

import java.math.BigDecimal;

public class SPSalesAmountBarCharVo {
    
	String shipDateMonth;
	BigDecimal bsAmount;
	BigDecimal odmPartsAmount;
	BigDecimal fgAmount;
	
	public String getShipDateMonth() {
		return shipDateMonth;
	}
	public void setShipDateMonth(String shipDateMonth) {
		this.shipDateMonth = shipDateMonth;
	}
	public BigDecimal getBsAmount() {
		return bsAmount;
	}
	public void setBsAmount(BigDecimal bsAmount) {
		this.bsAmount = bsAmount;
	}
	public BigDecimal getOdmPartsAmount() {
		return odmPartsAmount;
	}
	public void setOdmPartsAmount(BigDecimal odmPartsAmount) {
		this.odmPartsAmount = odmPartsAmount;
	}
	public BigDecimal getFgAmount() {
		return fgAmount;
	}
	public void setFgAmount(BigDecimal fgAmount) {
		this.fgAmount = fgAmount;
	}
	
}
