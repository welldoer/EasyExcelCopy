package net.blogjava.easyexcelcopy.metadata;

import org.apache.poi.ss.usermodel.IndexedColors;

public class TableStyle {
	private IndexedColors tableHeadBackgroundColor;
	private Font tableHeadFont;
	private Font tableContentFont;
	private IndexedColors tableContentBackgroundColor;

	public IndexedColors getTableHeadBackgroundColor() {
		return tableHeadBackgroundColor;
	}
	public void setTableHeadBackgroundColor(IndexedColors tableHeadBackgroundColor) {
		this.tableHeadBackgroundColor = tableHeadBackgroundColor;
	}
	public Font getTableHeadFont() {
		return tableHeadFont;
	}
	public void setTableHeadFont(Font tableHeadFont) {
		this.tableHeadFont = tableHeadFont;
	}
	public Font getTableContentFont() {
		return tableContentFont;
	}
	public void setTableContentFont(Font tableContentFont) {
		this.tableContentFont = tableContentFont;
	}
	public IndexedColors getTableContentBackgroundColor() {
		return tableContentBackgroundColor;
	}
	public void setTableContentBackgroundColor(IndexedColors tableContentBackgroundColor) {
		this.tableContentBackgroundColor = tableContentBackgroundColor;
	}
	
}
