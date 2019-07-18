package net.blogjava.easyexcelcopy;

import java.io.InputStream;
import java.util.List;

import net.blogjava.easyexcelcopy.analysis.ExcelAnalyserImpl;
import net.blogjava.easyexcelcopy.analysis.IExcelAnalyser;
import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.metadata.Sheet;
import net.blogjava.easyexcelcopy.support.ExcelTypeEnum;

public class ExcelReader {
    private IExcelAnalyser analyser = new ExcelAnalyserImpl();

    /**
     * @param in            文件输入流
     * @param excelTypeEnum excel类型03、07
     * @param customContent 自定义模型可以在{@link AnalysisEventListener#invoke(Object, AnalysisContext)
     *                      }AnalysisContext中获取用于监听者回调使用
     * @param eventListener 用户监听
     */
    public ExcelReader(InputStream in, ExcelTypeEnum excelTypeEnum, Object customContent,
                       AnalysisEventListener<?> eventListener) {
        this(in, excelTypeEnum, customContent, eventListener, true);
    }

    /**
     * @param in            文件输入流
     * @param excelTypeEnum excel类型03、07
     * @param customContent 自定义模型可以在{@link AnalysisEventListener#invoke(Object, AnalysisContext)
     *                      }AnalysisContext中获取用于监听者回调使用
     * @param eventListener 用户监听
     * @param trim          是否对解析的String做trim()默认true,用于防止 excel中空格引起的装换报错。
     */
    public ExcelReader(InputStream in, ExcelTypeEnum excelTypeEnum, Object customContent,
                       AnalysisEventListener<?> eventListener, boolean trim) {
        validateParam(in, excelTypeEnum, eventListener);
        analyser.init(in, excelTypeEnum, customContent, eventListener, trim);
    }

    /**
     * 读一个sheet，且没有模型映射
     */
    public void read() {
        analyser.analysis();
    }

    /**
     * 读指定个sheet，没有模型映射
     *
     * @param sheet 需要解析的sheet
     */
    public void read(Sheet sheet) {
        analyser.analysis(sheet);
    }

    /**
     * 读取excel中包含哪些sheet
     *
     * @return
     */
    public List<Sheet> getSheets() {
        return analyser.getSheets();
    }

    /**
     * 关闭流，删除临时目录文件
     */
    public void finish(){
        analyser.stop();
    }

    /**
     * 校验入参
     *
     * @param in
     * @param excelTypeEnum
     * @param eventListener
     */
    private void validateParam(InputStream in, ExcelTypeEnum excelTypeEnum, AnalysisEventListener eventListener) {
        if (eventListener == null) {
            throw new IllegalArgumentException("AnalysisEventListener can not null");
        } else if (in == null) {
            throw new IllegalArgumentException("InputStream can not null");
        } else if (excelTypeEnum == null) {
            throw new IllegalArgumentException("excelTypeEnum can not null");
        }
    }
}
