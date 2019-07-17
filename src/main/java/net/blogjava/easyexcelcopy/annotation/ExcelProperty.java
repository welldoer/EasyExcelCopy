package net.blogjava.easyexcelcopy.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

	/*
	 * 列所在的表头值
	 */
	String[] value() default {""};
	
	/*
	 * 列对应的顺序编号，越小越靠前
	 */
	int index() default 99999;
	
	String format() default "";
}
