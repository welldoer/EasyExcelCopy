package net.blogjava.easyexcelcopy.exception;

public class ExcelAnalysisException extends RuntimeException {
	private static final long serialVersionUID = -395300001685943672L;

	public ExcelAnalysisException() {
	}
	
	public ExcelAnalysisException(String message) {
		super(message);
	}

	public ExcelAnalysisException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelAnalysisException(Throwable cause) {
		super(cause);
	}
}
