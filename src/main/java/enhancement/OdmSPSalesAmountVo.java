package enhancement;

import java.math.BigDecimal;

public class OdmSPSalesAmountVo {
    
	String shipDateMonth;
	BigDecimal nbAmount;
	BigDecimal dtAmount;
	BigDecimal svrAmount;
	BigDecimal mntAmount;
	BigDecimal prjAmount;
	BigDecimal totalAmount;
	BigDecimal percent;
	
	public String getShipDateMonth() {
		return shipDateMonth;
	}
	public void setShipDateMonth(String shipDateMonth) {
		this.shipDateMonth = shipDateMonth;
	}
	public BigDecimal getNbAmount() {
		return nbAmount;
	}
	public void setNbAmount(BigDecimal nbAmount) {
		this.nbAmount = nbAmount;
	}
	public BigDecimal getDtAmount() {
		return dtAmount;
	}
	public void setDtAmount(BigDecimal dtAmount) {
		this.dtAmount = dtAmount;
	}
	public BigDecimal getSvrAmount() {
		return svrAmount;
	}
	public void setSvrAmount(BigDecimal svrAmount) {
		this.svrAmount = svrAmount;
	}
	public BigDecimal getMntAmount() {
		return mntAmount;
	}
	public void setMntAmount(BigDecimal mntAmount) {
		this.mntAmount = mntAmount;
	}
	public BigDecimal getPrjAmount() {
		return prjAmount;
	}
	public void setPrjAmount(BigDecimal prjAmount) {
		this.prjAmount = prjAmount;
	}
	public BigDecimal getTotalAmount() {
		return totalAmount;
	}
	public void setTotalAmount(BigDecimal totalAmount) {
		this.totalAmount = totalAmount;
	}
	public BigDecimal getPercent() {
		return percent;
	}
	public void setPercent(BigDecimal percent) {
		this.percent = percent;
	}
}
