package net.blogjava.easyexcelcopy.event;

import net.blogjava.easyexcelcopy.context.IAnalysisContext;

/*
 * 监听 Excel 解析每行数据，
 */
public abstract class AnalysisEventListener<T> {
	public abstract void invoke(T object, IAnalysisContext context);
	public abstract void doAfterAllAnalysed(IAnalysisContext context);
}
