package net.blogjava.easyexcelcopy.metadata;

import java.util.List;

public class Table {
	private Class<? extends BaseRowModel> clazz;
	// 当表头模型 clazz 不确定时，动态生成表头
	private List<List<String>> head;
	private int tableNo;
	private TableStyle tableStyle;
	
	public Table(int tableNo) {
		this.tableNo = tableNo;
	}

	public Class<? extends BaseRowModel> getClazz() {
		return clazz;
	}
	public void setClazz(Class<? extends BaseRowModel> clazz) {
		this.clazz = clazz;
	}
	public List<List<String>> getHead() {
		return head;
	}
	public void setHead(List<List<String>> head) {
		this.head = head;
	}
	public int getTableNo() {
		return tableNo;
	}
	public void setTableNo(int tableNo) {
		this.tableNo = tableNo;
	}
	public TableStyle getTableStyle() {
		return tableStyle;
	}
	public void setTableStyle(TableStyle tableStyle) {
		this.tableStyle = tableStyle;
	}
	
}
