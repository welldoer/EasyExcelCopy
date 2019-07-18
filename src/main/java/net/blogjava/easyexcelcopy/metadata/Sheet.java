package net.blogjava.easyexcelcopy.metadata;

import java.util.List;

import org.apache.commons.lang3.builder.ToStringBuilder;

public class Sheet {
	private int headlineNum;
	private int sheetNo;
	private String sheetName;
	private Class<? extends BaseRowModel> clazz;
	private List<List<String>> head;
	private TableStyle tableStyle;

    public Sheet(int sheetNo) {
        this.sheetNo = sheetNo;
    }
    public Sheet(int sheetNo, int headlineNum) {
        this.sheetNo = sheetNo;
        this.headlineNum = headlineNum;
    }
    public Sheet(int sheetNo, int headlineNum, Class<? extends BaseRowModel> clazz) {
        this.sheetNo = sheetNo;
        this.headlineNum = headlineNum;
        this.clazz = clazz;
    }
    public Sheet(int sheetNo, int headlineNum, Class<? extends BaseRowModel> clazz, String sheetName,
                 List<List<String>> head) {
        this.sheetNo = sheetNo;
        this.clazz = clazz;
        this.headlineNum = headlineNum;
        this.sheetName = sheetName;
        this.head = head;
    }
	
	public int getHeadlineNum() {
		return headlineNum;
	}
	public void setHeadlineNum(int headlineNum) {
		this.headlineNum = headlineNum;
	}
	public int getSheetNo() {
		return sheetNo;
	}
	public void setSheetNo(int sheetNo) {
		this.sheetNo = sheetNo;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public Class<? extends BaseRowModel> getClazz() {
		return clazz;
	}
	public void setClazz(Class<? extends BaseRowModel> clazz) {
		this.clazz = clazz;
        if (headlineNum == 0) {
            this.headlineNum = 1;
        }
	}
	public List<List<String>> getHead() {
		return head;
	}
	public void setHead(List<List<String>> head) {
		this.head = head;
	}
	public TableStyle getTableStyle() {
		return tableStyle;
	}
	public void setTableStyle(TableStyle tableStyle) {
		this.tableStyle = tableStyle;
	}

	@Override
	public String toString() {
		return ToStringBuilder.reflectionToString(this);
	}
}
