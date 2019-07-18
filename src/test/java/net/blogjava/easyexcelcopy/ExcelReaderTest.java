package net.blogjava.easyexcelcopy;

import static org.assertj.core.api.Assertions.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.List;

import org.junit.Before;
import org.junit.Test;

import net.blogjava.easyexcelcopy.annotation.ExcelProperty;
import net.blogjava.easyexcelcopy.context.IAnalysisContext;
import net.blogjava.easyexcelcopy.event.AnalysisEventListener;
import net.blogjava.easyexcelcopy.metadata.BaseRowModel;
import net.blogjava.easyexcelcopy.metadata.Sheet;
import net.blogjava.easyexcelcopy.support.ExcelTypeEnum;

public class ExcelReaderTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void testInterrupt() {
		InputStream inputStream = getInputStream("2007WithModelMultipleSheet.xlsx");
        try {
            ExcelReader reader = new ExcelReader(inputStream, ExcelTypeEnum.XLSX, null,
                new AnalysisEventListener<Object>() {
                    @Override
                    public void invoke(Object object, IAnalysisContext context) {
                        context.interrupt();
                    }

                    @Override
                    public void doAfterAllAnalysed(IAnalysisContext context) {
                    }
                });

            List<Sheet> sheets = reader.getSheets();
            System.out.println(sheets);
            for (Sheet sheet : sheets) {
                if (sheet.getSheetNo() == 1) {
                    sheet.setHeadlineNum(2);
                    sheet.setClazz(ExcelRowJavaModel.class);
                }
                if (sheet.getSheetNo() == 2) {
                    sheet.setHeadlineNum(1);
                    sheet.setClazz(ExcelRowJavaModel1.class);
                }
                reader.read(sheet);
            }

        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
	}

	private InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }
}

class ExcelRowJavaModel extends BaseRowModel {
    @ExcelProperty(index = 0,value = "银行放款编号")
    private int num;

    @ExcelProperty(index = 1,value = "code")
    private Long code;

    @ExcelProperty(index = 2,value = "银行存放期期")
    private Date endTime;

    @ExcelProperty(index = 3,value = "测试1")
    private Double money;

    @ExcelProperty(index = 4,value = "测试2")
    private String times;

    @ExcelProperty(index = 5,value = "测试3")
    private int activityCode;

    @ExcelProperty(index = 6,value = "测试4")
    private Date date;

    @ExcelProperty(index = 7,value = "测试5")
    private Double lx;

    @ExcelProperty(index = 8,value = "测试6")
    private String name;

    public int getNum() {
        return num;
    }

    public void setNum(int num) {
        this.num = num;
    }

    public Long getCode() {
        return code;
    }

    public void setCode(Long code) {
        this.code = code;
    }

    public Date getEndTime() {
        return endTime;
    }

    public void setEndTime(Date endTime) {
        this.endTime = endTime;
    }

    public Double getMoney() {
        return money;
    }

    public void setMoney(Double money) {
        this.money = money;
    }

    public String getTimes() {
        return times;
    }

    public void setTimes(String times) {
        this.times = times;
    }

    public int getActivityCode() {
        return activityCode;
    }

    public void setActivityCode(int activityCode) {
        this.activityCode = activityCode;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Double getLx() {
        return lx;
    }

    public void setLx(Double lx) {
        this.lx = lx;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "ExcelRowJavaModel{" +
            "num=" + num +
            ", code=" + code +
            ", endTime=" + endTime +
            ", money=" + money +
            ", times='" + times + '\'' +
            ", activityCode=" + activityCode +
            ", date=" + date +
            ", lx=" + lx +
            ", name='" + name + '\'' +
            '}';
    }
}

class ExcelRowJavaModel1 extends BaseRowModel {

    @ExcelProperty(index = 0,value = "银行放款编号")
    private int num;

    @ExcelProperty(index = 1,value = "code")
    private Long code;

    @ExcelProperty(index = 2,value = "银行存放期期")
    private Date endTime;

    @ExcelProperty(index = 3,value = "测试1")
    private Double money;

    @ExcelProperty(index = 4,value = "测试2")
    private String times;

    @ExcelProperty(index = 5,value = "测试3")
    private int activityCode;

    @ExcelProperty(index = 6,value = "测试4")
    private Date date;

    @ExcelProperty(index = 7,value = "测试5")
    private Double lx;

    @ExcelProperty(index = 8,value = "测试6")
    private String name;


    public int getNum() {
        return num;
    }

    public void setNum(int num) {
        this.num = num;
    }

    public Long getCode() {
        return code;
    }

    public void setCode(Long code) {
        this.code = code;
    }

    public Date getEndTime() {
        return endTime;
    }

    public void setEndTime(Date endTime) {
        this.endTime = endTime;
    }

    public Double getMoney() {
        return money;
    }

    public void setMoney(Double money) {
        this.money = money;
    }

    public String getTimes() {
        return times;
    }

    public void setTimes(String times) {
        this.times = times;
    }

    public int getActivityCode() {
        return activityCode;
    }

    public void setActivityCode(int activityCode) {
        this.activityCode = activityCode;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Double getLx() {
        return lx;
    }

    public void setLx(Double lx) {
        this.lx = lx;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "ExcelRowJavaModel{" +
            "num=" + num +
            ", code=" + code +
            ", endTime=" + endTime +
            ", money=" + money +
            ", times='" + times + '\'' +
            ", activityCode=" + activityCode +
            ", date=" + date +
            ", lx=" + lx +
            ", name='" + name + '\'' +
            '}';
    }
}
