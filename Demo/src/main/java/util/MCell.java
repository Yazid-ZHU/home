package util;

public class MCell {
	
	private int cellNum;
	
	private String value;
	
	public MCell() {
		super();
	}

	public MCell(int cellNum, String value) {
		super();
		this.cellNum = cellNum;
		this.value = value;
	}

	public int getCellNum() {
		return cellNum;
	}

	public void setCellNum(int cellNum) {
		this.cellNum = cellNum;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	@Override
	public String toString() {
		return "MCell [cellNum=" + cellNum + ", value=" + value + "]";
	}
	
}
