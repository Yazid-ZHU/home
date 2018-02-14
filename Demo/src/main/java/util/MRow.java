package util;

import java.util.*;

public class MRow {
	
	private Integer rowNum;
	
	private Map<String, MCell> cells = new LinkedHashMap<String, MCell>();

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(Integer rowNum) {
		this.rowNum = rowNum;
	}

	public Map<String, MCell> getCells() {
		return cells;
	}

	public void setCells(Map<String, MCell> cells) {
		this.cells = cells;
	}

	@Override
	public String toString() {
		return "MRow [rowNum=" + rowNum + ", cells=" + cells + "]";
	}

}
