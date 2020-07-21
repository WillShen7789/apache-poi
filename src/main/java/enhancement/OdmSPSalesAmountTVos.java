package enhancement;

import java.math.BigDecimal;
import java.util.List;

public class OdmSPSalesAmountTVos {
		
	private List<String> shipDateMonth;

	private int rowSize = 0;
	
	public List<String> getShipDateMonth() {
		return shipDateMonth;
	}
	public void setShipDateMonth(List<String> shipDateMonth) {
		this.shipDateMonth = shipDateMonth;
	}
	public int getRowSize() {
		return rowSize;
	}
	public void setRowSize(int rowSize) {
		this.rowSize = rowSize;
	}
	
}
