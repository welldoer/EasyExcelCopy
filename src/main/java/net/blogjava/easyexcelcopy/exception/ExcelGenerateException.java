package net.blogjava.easyexcelcopy.exception;

public class ExcelGenerateException extends RuntimeException {
	private static final long serialVersionUID = 4795209105240514424L;

	public ExcelGenerateException(String message) {
        super(message);
    }

    public ExcelGenerateException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelGenerateException(Throwable cause) {
        super(cause);
    }
}
