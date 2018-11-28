package cn.excel;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

/**
 * 
 * @author yutyi
 *
 */
@Retention(RUNTIME)
@Target({ FIELD })
public @interface ExcelField {
	
	/**
	 * 是否检查必填
	 * @return
	 */
	boolean required() default false;
	
	/**
	 * 是否检查唯一
	 * @return
	 */
	boolean unique() default false;
	
	/**
	 * 检查是否指定格式<br/>
	 * string\date\datetime\time\int\double\boolean\regex
	 * @return
	 */
	String format() default "";
	
	/**
	 * 属性处理器
	 * @return
	 */
	Class<?> handleField() default Class.class;
	
	/**
	 * 字段排序号
	 * @return
	 */
	int sort() default 0;

	/**
	 * 导出Excel列名
	 */
	String columnName() default "";
}
