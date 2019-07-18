package net.blogjava.easyexcelcopy.analysis;

import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import net.blogjava.easyexcelcopy.context.IAnalysisContext;
import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.event.IAnalysisEventRegisterCenter;
import net.blogjava.easyexcelcopy.event.OneRowAnalysisFinishEvent;
import net.blogjava.easyexcelcopy.metadata.Sheet;
import net.blogjava.easyexcelcopy.support.ExcelTypeEnum;

public abstract class BaseSaxAnalyser implements IAnalysisEventRegisterCenter, IExcelAnalyser {
	protected IAnalysisContext analysisContext;
	private Map<String, AnalysisEventListener> listeners = new LinkedHashMap<>();

	protected abstract void execute();
	
	@Override
	public void init(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
			AnalysisEventListener<?> eventListener, boolean trim) {
	}

	@Override
	public void analysis(Sheet sheetParam) {
		execute();
	}

	@Override
	public void analysis() {
		execute();
	}

	@Override
	public List<Sheet> getSheets() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void stop() {
		// TODO Auto-generated method stub

	}

	@Override
	public void appendLister(String name, AnalysisEventListener<?> listener) {
		if (!listeners.containsKey(name)) {
            listeners.put(name, listener);
        }
	}

	@Override
	public void notifyListeners(OneRowAnalysisFinishEvent event) {
		analysisContext.setCurrentRowAnalysisResult(event.getData());

        //表头数据
        if (analysisContext.getCurrentRowNum() < analysisContext.getCurrentSheet().getHeadlineNum()) {
            if (analysisContext.getCurrentRowNum() <= analysisContext.getCurrentSheet().getHeadlineNum() - 1) {
                analysisContext.buildExcelHeadProperty(null,
                    (List<String>)analysisContext.getCurrentRowAnalysisResult());
            }
        } else {
            analysisContext.setCurrentRowAnalysisResult(event.getData());
            for (Map.Entry<String, AnalysisEventListener> entry : listeners.entrySet()) {
                entry.getValue().invoke(analysisContext.getCurrentRowAnalysisResult(), analysisContext);
            }
        }
	}

	@Override
	public void cleanAllListeners() {
		listeners = new LinkedHashMap<String, AnalysisEventListener>();
	}

}
