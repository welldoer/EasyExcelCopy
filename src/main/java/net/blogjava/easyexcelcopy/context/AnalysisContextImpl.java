package net.blogjava.easyexcelcopy.context;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.exception.ExcelAnalysisException;
import net.blogjava.easyexcelcopy.metadata.BaseRowModel;
import net.blogjava.easyexcelcopy.metadata.ExcelHeadProperty;
import net.blogjava.easyexcelcopy.metadata.Sheet;
import net.blogjava.easyexcelcopy.support.ExcelTypeEnum;

// 提供解析 Excel 上线文默认实现
public class AnalysisContextImpl implements IAnalysisContext {
	private Object custom;
	private Sheet currentSheet;
	private ExcelTypeEnum excelType;
	private InputStream inputStream;
	private AnalysisEventListener<?> eventListener;
	private int currentRowNum;
	private int totalCount;
	private ExcelHeadProperty excelHeadProperty;
	private boolean trim;
	private boolean use1904WindowDate = false;
	private Object currentRowAnalysisResult;

	public AnalysisContextImpl(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
			AnalysisEventListener<?> listener, boolean trim) {
		this.custom = custom;
		this.eventListener = listener;
		this.inputStream = inputStream;
		this.excelType = excelTypeEnum;
		this.trim = trim;
	}
	
	@Override
	public Object getCustom() {
		return custom;
	}

	@Override
	public Sheet getCurrentSheet() {
		return currentSheet;
	}

	@Override
	public void setCurrentSheet(Sheet sheet) {
		this.currentSheet = sheet;
		if (currentSheet.getClazz() != null) {
            buildExcelHeadProperty(currentSheet.getClazz(), null);
        }
	}

	@Override
	public ExcelTypeEnum getExcelType() {
		return excelType;
	}

	@Override
	public InputStream getInputStream() {
		return inputStream;
	}

	@Override
	public AnalysisEventListener<?> getEventListener() {
		return eventListener;
	}

	@Override
	public int getCurrentRowNum() {
		return currentRowNum;
	}

	@Override
	public void setCurrentRowNum(int rowNum) {
		this.currentRowNum = rowNum;
	}

	@Override
	public int getTotalCount() {
		return totalCount;
	}

	@Override
	public void setTotalCount(int totalCount) {
		this.totalCount = totalCount;
	}

	@Override
	public ExcelHeadProperty getExcelHeadProperty() {
		return excelHeadProperty;
	}

	@Override
	public void buildExcelHeadProperty(Class<? extends BaseRowModel> clazz, List<String> headOneRow) {
		if (this.excelHeadProperty == null && (clazz != null || headOneRow != null)) {
            this.excelHeadProperty = new ExcelHeadProperty(clazz, new ArrayList<List<String>>());
        }
        if (this.excelHeadProperty.getHead() == null && headOneRow != null) {
            this.excelHeadProperty.appendOneRow(headOneRow);
        }
	}

	@Override
	public boolean trim() {
		return trim;
	}

	@Override
	public void setCurrentRowAnalysisResult(Object result) {
		this.currentRowAnalysisResult = result;
	}

	@Override
	public Object getCurrentRowAnalysisResult() {
		return currentRowAnalysisResult;
	}

	@Override
	public void interrupt() {
		throw new ExcelAnalysisException("interrupt error");
	}

	@Override
	public boolean use1904WindowDate() {
		return use1904WindowDate;
	}

	@Override
	public void setUse1904WindowDate(boolean use1904WindowDate) {
		this.use1904WindowDate = use1904WindowDate;
	}

}
