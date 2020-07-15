package enhancement;

import java.math.BigDecimal;
import java.util.List;

public class OdmSPSalesAmountTVos {
		
	private List<String> shipDateMonth;
	private List<BigDecimal> nbAmount;
	private List<BigDecimal> dtAmount;
	private List<BigDecimal> svrAmount;
	private List<BigDecimal> mntAmount;
	private List<BigDecimal> prjAmount;
	private List<BigDecimal> totalAmount;
	private List<BigDecimal> percent;

	private int size = 0;
	
	public List<String> getShipDateMonth() {
		return shipDateMonth;
	}
	public void setShipDateMonth(List<String> shipDateMonth) {
		this.shipDateMonth = shipDateMonth;
	}
	public List<BigDecimal> getNbAmount() {
		return nbAmount;
	}
	public void setNbAmount(List<BigDecimal> nbAmount) {
		this.nbAmount = nbAmount;
	}
	public List<BigDecimal> getDtAmount() {
		return dtAmount;
	}
	public void setDtAmount(List<BigDecimal> dtAmount) {
		this.dtAmount = dtAmount;
	}
	public List<BigDecimal> getSvrAmount() {
		return svrAmount;
	}
	public void setSvrAmount(List<BigDecimal> svrAmount) {
		this.svrAmount = svrAmount;
	}
	public List<BigDecimal> getMntAmount() {
		return mntAmount;
	}
	public void setMntAmount(List<BigDecimal> mntAmount) {
		this.mntAmount = mntAmount;
	}
	public List<BigDecimal> getPrjAmount() {
		return prjAmount;
	}
	public void setPrjAmount(List<BigDecimal> prjAmount) {
		this.prjAmount = prjAmount;
	}
	public List<BigDecimal> getTotalAmount() {
		return totalAmount;
	}
	public void setTotalAmount(List<BigDecimal> totalAmount) {
		this.totalAmount = totalAmount;
	}
	public List<BigDecimal> getPercent() {
		return percent;
	}
	public void setPercent(List<BigDecimal> percent) {
		this.percent = percent;
	}
	public int getSize() {
		return size;
	}
	public void setSize(int size) {
		this.size = size;
	}
}
