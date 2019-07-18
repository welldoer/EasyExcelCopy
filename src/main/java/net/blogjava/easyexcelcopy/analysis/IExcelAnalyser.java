package net.blogjava.easyexcelcopy.analysis;

import java.io.InputStream;
import java.util.List;

import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.metadata.Sheet;
import net.blogjava.easyexcelcopy.support.ExcelTypeEnum;

public interface IExcelAnalyser {
	void init(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom, AnalysisEventListener<?> eventListener,
            boolean trim);
	void analysis(Sheet sheetParam);
	void analysis();
	List<Sheet> getSheets();
	void stop();
}
