package cn.excel;

import java.util.Map;

/**
 * 属性处理器
 * 
 * @author yutyi
 *
 */
public abstract class HandleField {
	
	public abstract void setParams(Object params);

	/**
	 * 获取错误消息
	 * 
	 * @return
	 */
	public abstract String getMessage();

	/**
	 * 验证数据格式
	 * 
	 * @param columns
	 * @return
	 */
	public abstract boolean validate(String fieldName,Map<String, Object> columns);

	/**
	 * 转换数据
	 * 
	 * @param columns
	 * @return
	 */
	public abstract Object translate(Map<String, Object> columns);

}
