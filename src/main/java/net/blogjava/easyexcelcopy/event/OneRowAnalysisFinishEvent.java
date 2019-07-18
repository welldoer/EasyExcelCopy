package net.blogjava.easyexcelcopy.event;

public class OneRowAnalysisFinishEvent {
	private Object data;

	public OneRowAnalysisFinishEvent(Object data) {
		this.data = data;
	}

	public Object getData() {
		return data;
	}
	public void setData(Object data) {
		this.data = data;
	}
}
