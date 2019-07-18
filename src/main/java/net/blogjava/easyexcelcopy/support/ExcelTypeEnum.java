package net.blogjava.easyexcelcopy.support;

public enum ExcelTypeEnum {
	XLS(".xls"),
	XLSX(".xlsx");
	
	private String value;
	
	private ExcelTypeEnum(String value) {
		this.value = value;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}
}
