package net.blogjava.easyexcelcopy.modelbuild;

import java.util.List;

import org.apache.commons.beanutils.BeanUtils;

import net.blogjava.easyexcelcopy.context.IAnalysisContext;
import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.exception.ExcelGenerateException;
import net.blogjava.easyexcelcopy.metadata.ExcelColumnProperty;
import net.blogjava.easyexcelcopy.metadata.ExcelHeadProperty;
import net.blogjava.easyexcelcopy.util.TypeUtil;

public class ModelBuildEventListener extends AnalysisEventListener {


    @Override
    public void invoke(Object object, IAnalysisContext context) {


        if(context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz()!=null ){
            Object resultModel = buildUserModel(context, (List<String>)object);
            context.setCurrentRowAnalysisResult(resultModel);
        }

    }



    private Object buildUserModel(IAnalysisContext context, List<String> stringList) {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();

        Object resultModel;
        try {
            resultModel = excelHeadProperty.getHeadClazz().newInstance();
        } catch (Exception e) {
            throw new ExcelGenerateException(e);
        }
        if (excelHeadProperty != null) {
            for (int i = 0; i < stringList.size(); i++) {
                ExcelColumnProperty columnProperty = excelHeadProperty.getExcelColumnProperty(i);
                if (columnProperty != null) {
                    Object value = TypeUtil.convert(stringList.get(i), columnProperty.getField(),
                        columnProperty.getFormat(),context.use1904WindowDate());
                    if (value != null) {
                        try {
                            BeanUtils.setProperty(resultModel, columnProperty.getField().getName(), value);
                        } catch (Exception e) {
                            throw new ExcelGenerateException(
                                columnProperty.getField().getName() + " can not set value " + value, e);
                        }
                    }
                }
            }
        }
        return resultModel;
    }

    @Override
    public void doAfterAllAnalysed(IAnalysisContext context) {

    }
}
