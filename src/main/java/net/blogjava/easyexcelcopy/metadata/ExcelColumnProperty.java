package net.blogjava.easyexcelcopy.metadata;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.builder.ToStringBuilder;

public class ExcelColumnProperty implements Comparable<ExcelColumnProperty> {
	private Field field;
	private int index = 99999;
	private List<String> head = new ArrayList<>();
	private String format;

	public Field getField() {
		return field;
	}
	public void setField(Field field) {
		this.field = field;
	}
	public int getIndex() {
		return index;
	}
	public void setIndex(int index) {
		this.index = index;
	}
	public List<String> getHead() {
		return head;
	}
	public void setHead(List<String> head) {
		this.head = head;
	}
	public String getFormat() {
		return format;
	}
	public void setFormat(String format) {
		this.format = format;
	}

	@Override
	public int compareTo(ExcelColumnProperty o) {
		return (index == o.index) ? 0 : ((index < o.index) ? -1 : 1);
	}

	@Override
	public String toString() {
		return ToStringBuilder.reflectionToString(this);
	}
}
