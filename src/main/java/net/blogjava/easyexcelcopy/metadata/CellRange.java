package net.blogjava.easyexcelcopy.metadata;

import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.ToStringBuilder;

public class CellRange {
	private int firstRow;
	private int lastRow;
	private int firstCol;
	private int lastCol;
	
	public CellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}
	
	public int getFirstRow() {
		return firstRow;
	}
	public void setFirstRow(int firstRow) {
		this.firstRow = firstRow;
	}
	public int getLastRow() {
		return lastRow;
	}
	public void setLastRow(int lastRow) {
		this.lastRow = lastRow;
	}
	public int getFirstCol() {
		return firstCol;
	}
	public void setFirstCol(int firstCol) {
		this.firstCol = firstCol;
	}
	public int getLastCol() {
		return lastCol;
	}
	public void setLastCol(int lastCol) {
		this.lastCol = lastCol;
	}

	@Override
	public boolean equals(Object obj) {
		return EqualsBuilder.reflectionEquals(this, obj);
	}
	
	@Override
	public String toString() {
		return ToStringBuilder.reflectionToString(this);
	}
}
