package net.blogjava.easyexcelcopy.annotation;

import static org.assertj.core.api.Assertions.*;

import java.util.Date;

import org.junit.Before;
import org.junit.Test;

public class ExcelPropertyTest {
	
	@ExcelProperty(index = 0, value = "银行放款编号")
	private int bankNum;
	
	@ExcelProperty(index = 2, value = "银行到期日", format = "yyyy/mm/dd")
	private Date endDate;

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void test() throws NoSuchFieldException, SecurityException {
		assertThat(bankNum).isEqualTo(0);
		assertThat(endDate).isNull();;
		
		ExcelProperty excelProperty = this.getClass().getDeclaredField("bankNum").getAnnotation(ExcelProperty.class);
		assertThat(excelProperty.value()).isEqualTo(new String[] {"银行放款编号"});
		assertThat(excelProperty.index()).isEqualTo(0);
		assertThat(excelProperty.format()).isEqualTo("");
		
		excelProperty = this.getClass().getDeclaredField("endDate").getAnnotation(ExcelProperty.class);
		assertThat(excelProperty.value()).isEqualTo(new String[] {"银行到期日"});
		assertThat(excelProperty.index()).isEqualTo(2);
		assertThat(excelProperty.format()).isEqualTo("yyyy/mm/dd");
	}

}
