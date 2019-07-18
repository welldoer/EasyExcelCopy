package net.blogjava.easyexcelcopy.event;

public interface IAnalysisEventRegisterCenter {
	void appendLister(String name, AnalysisEventListener<?> listener);
	void notifyListeners(OneRowAnalysisFinishEvent event);
	void cleanAllListeners();
}
