package net.blogjava.easyexcelcopy.metadata;

import static org.assertj.core.api.Assertions.*;

import java.util.Date;

import org.junit.Before;
import org.junit.Test;

import net.blogjava.easyexcelcopy.annotation.ExcelColumnNum;
import net.blogjava.easyexcelcopy.annotation.ExcelProperty;

public class ExcelHeadPropertyTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void testNewBasicExcelHeadProperty() {
		BaseRowModel headClazz = new BaseRowModel();
		ExcelHeadProperty excelHeadProperty = new ExcelHeadProperty(headClazz.getClass(), null);
		
		assertThat(excelHeadProperty.getHeadClazz()).isEqualTo(BaseRowModel.class);
		assertThat(excelHeadProperty.getRowNum()).isEqualTo(0);
	}

	@Test
	public void testTwoColumnsExcelHeadProperty() {
		BaseRowModel headClazz = new ThreeColumnHeadClass();
		ExcelHeadProperty excelHeadProperty = new ExcelHeadProperty(headClazz.getClass(), null);
		
		assertThat(excelHeadProperty.getHeadClazz()).isEqualTo(ThreeColumnHeadClass.class);
		assertThat(excelHeadProperty.getRowNum()).isEqualTo(2);
		assertThat(excelHeadProperty.getHead().size()).isEqualTo(3);
		assertThat(excelHeadProperty.getHead().get(0).get(0)).isEqualTo("第二列");
		assertThat(excelHeadProperty.getHead().get(1).get(0)).isEqualTo("第一列");
		assertThat(excelHeadProperty.getHead().get(1).get(1)).isEqualTo("首列");
		assertThat(excelHeadProperty.getHead().get(2).get(0)).isEqualTo("第三列");
		assertThat(excelHeadProperty.getColumnPropertyList().size()).isEqualTo(3);
		assertThat(excelHeadProperty.getExcelColumnProperty(0).getIndex()).isEqualTo(2);
		assertThat(excelHeadProperty.getExcelColumnProperty(0).getHead().get(0)).isEqualTo("第二列");
		assertThat(excelHeadProperty.getExcelColumnProperty(1).getIndex()).isEqualTo(99999);
		assertThat(excelHeadProperty.getExcelColumnProperty(1).getHead().get(0)).isEqualTo("第一列");
		assertThat(excelHeadProperty.getExcelColumnProperty(1).getHead().get(1)).isEqualTo("首列");
		assertThat(excelHeadProperty.getExcelColumnProperty(2).getIndex()).isEqualTo(2);
//		assertThat(excelHeadProperty.getExcelColumnProperty(2).getHead().get(0)).isEqualTo("第三列");
		assertThat(excelHeadProperty.getCellRangeModels().size()).isEqualTo(2);
		assertThat(excelHeadProperty.getCellRangeModels().get(0)).isEqualTo(new CellRange(0, 1, 0, 0));
//		assertThat(excelHeadProperty.getCellRangeModels().get(1)).isEqualTo(new CellRange(0, 1, 2, 2));
	}

	class ThreeColumnHeadClass extends BaseRowModel {
		@ExcelProperty(value = {"第一列","首列"})
		private String firstColumn;

		@ExcelProperty(index = 2, value = "第二列")
		private int secondColumn;

		@ExcelProperty(value = "第三列", format = "yyyy-mm-dd")
		private Date endDate;
	}

	@Test
	public void testSimpleColumnsExcelHeadProperty() {
		BaseRowModel headClazz = new SimpleColumnsHeadClass();
		ExcelHeadProperty excelHeadProperty = new ExcelHeadProperty(headClazz.getClass(), null);
		
		assertThat(excelHeadProperty.getHeadClazz()).isEqualTo(SimpleColumnsHeadClass.class);
		assertThat(excelHeadProperty.getRowNum()).isEqualTo(0);
		assertThat(excelHeadProperty.getHead().size()).isEqualTo(2);
		assertThat(excelHeadProperty.getHead().get(0).size()).isEqualTo(0);
		assertThat(excelHeadProperty.getHead().get(1).size()).isEqualTo(0);
		assertThat(excelHeadProperty.getColumnPropertyList().size()).isEqualTo(2);
		assertThat(excelHeadProperty.getExcelColumnProperty(0).getIndex()).isEqualTo(2);
	}

	class SimpleColumnsHeadClass extends BaseRowModel {
		@ExcelColumnNum(value = 3, format = "yyyy/mm/dd")
		private Date beginDate;

		@ExcelColumnNum(value = 2)
		private int secondColumn;
	}
}
