package net.blogjava.easyexcelcopy.context;

import java.io.InputStream;
import java.util.List;

import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.metadata.BaseRowModel;
import net.blogjava.easyexcelcopy.metadata.ExcelHeadProperty;
import net.blogjava.easyexcelcopy.metadata.Sheet;
import net.blogjava.easyexcelcopy.support.ExcelTypeEnum;

public interface IAnalysisContext {
	// 返回用户自定义数据
	Object getCustom();
	Sheet getCurrentSheet();
	void setCurrentSheet(Sheet sheet);
	ExcelTypeEnum getExcelType();
	InputStream getInputStream();
	AnalysisEventListener<?> getEventListener();
	int getCurrentRowNum();
	void setCurrentRowNum(int rowNum);
	@Deprecated
	int getTotalCount();
	void setTotalCount(int totalCount);
	ExcelHeadProperty getExcelHeadProperty();
	void buildExcelHeadProperty(Class<? extends BaseRowModel> clazz, List<String> headOneRow);
	boolean trim();
	void setCurrentRowAnalysisResult(Object result);
	Object getCurrentRowAnalysisResult();
	void interrupt();
	boolean  use1904WindowDate();
	void setUse1904WindowDate(boolean use1904WindowDate);
}
