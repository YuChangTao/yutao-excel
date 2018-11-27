package cn.excel.imp;

import cn.excel.ExcelField;
import cn.excel.util.DateUtils;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang.ObjectUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

/**
 * Excel表格导入数据库
 *
 * @author yutyi
 */
public class ExcelImport {

	private final static Logger logger = LoggerFactory.getLogger(ExcelImport.class);

	private Class<?> clazz;

	/**
	 * 文件路径
	 */
	private String filepath;

	/**
	 * 列注解列表
	 */
	private List<Object[]> annotationList;

	/**
	 * 
	 */
	private List<Map<String, Object>> dataList;

	/**
	 * 工作薄对象
	 */
	private Workbook wb;

	/**
	 * 工作表对象
	 */
	private Sheet sheet;

	/**
	 * 公式计算器
	 */
	private FormulaEvaluator evaluator;

	/**
	 * 画板
	 */
	private Drawing<?> patriarch;

	/**
	 * 
	 */
	private CreationHelper creationHelper;

	private SimpleDateFormat sdf14 = new SimpleDateFormat("yyyy/M/d");

	private SimpleDateFormat sdf22 = new SimpleDateFormat("yyyy/M/d HH:mm");

	private SimpleDateFormat sdf177 = new SimpleDateFormat("yyyy年M月d日");

	private NumberFormat numberFormat = NumberFormat.getInstance();

	/**
	 * 工作表中行数
	 */
	private int rowCount;

	/**
	 * 数据起始索引
	 */
	private int dataIndex;

	/**
	 * 传递的参数
	 */
	private Object params;

	public ExcelImport(Class<?> clazz, String filepath) {
		this(clazz, 1, filepath);
	}

	/**
	 * 
	 * @param clazz
	 * @param dataIndex
	 *            数据列从1开始
	 * @param filepath
	 */
	public ExcelImport(Class<?> clazz, int dataIndex, String filepath) {
		this(clazz, dataIndex, null, filepath);
	}

	public ExcelImport(Class<?> clazz, int dataIndex, Object params, String filePath) {
		logger.info("==>  Excel文件地址：{}", filePath);
		this.clazz = clazz;
		this.dataIndex = dataIndex;
		this.params = params;
		this.filepath = filePath;
		numberFormat.setGroupingUsed(false);
		this.init();
	}

	public void init() {
		// 检查文件是否靠谱
		if (StringUtils.isBlank(filepath)) {
			throw new RuntimeException("导入文档为空!");
		} else if (filepath.toLowerCase().endsWith(".xls") || filepath.toLowerCase().endsWith(".xlsx")) {
			try {
				InputStream is = new FileInputStream(filepath);
				wb = WorkbookFactory.create(is);
				creationHelper = wb.getCreationHelper();
				evaluator = creationHelper.createFormulaEvaluator();
			} catch (Exception ex) {
				throw new RuntimeException("文档格式不正确!");
			}
		} else {
			throw new RuntimeException("文档格式不正确!");
		}
		// 获取工作薄
		this.sheet = this.wb.getSheetAt(0);
		this.patriarch = this.sheet.createDrawingPatriarch();

		this.rowCount = this.sheet.getLastRowNum();
		if (this.rowCount == 0) {
			throw new RuntimeException("模板格式错误，第一个工作薄无列头");
		}
		// 获取列上的注解，并根据注解进行排序
		List<Object[]> annotationList = new ArrayList<>();
		Field[] fields = this.clazz.getDeclaredFields();
		for (Field field : fields) {
			ExcelField excelField = field.getAnnotation(ExcelField.class);
			if (excelField != null) {
				annotationList.add(new Object[] { excelField, field });
			}
		}
		// Field sorting
		Collections.sort(annotationList, new Comparator<Object[]>() {
			public int compare(Object[] o1, Object[] o2) {
				return new Integer(((ExcelField) o1[0]).sort()).compareTo(new Integer(((ExcelField) o2[0]).sort()));
			};
		});
		this.annotationList = annotationList;

		// 读取文件数据
		this.dataList = this.getDataList();
		// this.dataNum = this.dataList.size();
	}

	private List<Map<String, Object>> getDataList() {
		List<Map<String, Object>> dataList = new ArrayList<>();
		int rowCount = this.sheet.getLastRowNum();
		int columnCount = this.annotationList.size();
		for (int i = dataIndex; i <= rowCount; i++) {
			Map<String, Object> map = new HashMap<>();
			Row row = this.getRow(i);
			row = row == null ? this.sheet.createRow(i) : row;
			for (int j = 0; j < columnCount; j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					continue;
				}
				String cellValue = this.getCellValue(cell);
				cellValue = StringUtils.trim(cellValue);
				if (StringUtils.isNotEmpty(cellValue)) {
					Object[] objs = this.annotationList.get(j);
					Field field = (Field) objs[1];
					String fieldName = field.getName();
					map.put(fieldName, cellValue);
				}
				// 删除批注
				cell.removeCellComment();
				this.clearError(cell);
			}
			dataList.add(map);
			this.rowCount += (map.isEmpty() ? 0 : 1);
		}
		return dataList;
	}

	private int dataNum; // 数据行数
	private int successNum; // 成功行数
	private int errorNum; // 错误行数

	public boolean validate() {
		int rowNum = this.dataList.size();

		boolean isError = true;
		for (int i = 0; i < rowNum; i++) {
			Row row = this.sheet.getRow(i + this.dataIndex);
			Map<String, Object> data = this.dataList.get(i);
			if (data != null && !data.isEmpty()) {
				if (validate(row, data)) {
					successNum++;
				} else {
					isError = false;
					errorNum++;
				}
			}
		}
		return isError;
	}

	/**
	 * 
	 * @param data
	 * @return
	 */
	private boolean validate(Row row, Map<String, Object> data) {
		boolean result = true;
		try {
			int columnNum = this.annotationList.size();
			for (int i = 0; i < columnNum; i++) {
				Object[] objs = annotationList.get(i);
				ExcelField excelField = (ExcelField) objs[0];
				Field field = (Field) objs[1];
				String fieldName = field.getName();

				String fieldValue = ObjectUtils.toString(data.get(fieldName));
				Cell cell = row.getCell(i);
				if (StringUtils.isEmpty(fieldValue)) {
					// 检查是否有值
					if (excelField.required()) {
						cell = (cell == null ? row.createCell(i) : cell); // 单元格不存在时，则创建一个
						this.flagError(cell, "不允许空");
						result = false;
					}
					continue;
				}
				// 检查是否唯一
				if (excelField.unique()) {
					List<String> valueList = uniqueMap.get(fieldName);
					valueList = (valueList == null ? new ArrayList<>() : valueList);
					if (valueList.contains(fieldValue)) {
						this.flagError(cell, "该列不允许重复");
						result = false;
						continue;
					}
					valueList.add(fieldValue);
					uniqueMap.put(fieldName, valueList);
				}
				// 检查格式是否正确
				if (!checkFormat(cell, field, excelField, fieldValue)) {
					result = false;
					continue;
				}
				// handleField检查是否正确
				if (!checkHandle(cell, field, excelField, data)) {
					result = false;
					continue;
				}
			}
		} catch (Exception ex) {
			logger.error("excel检查报错", ex);
		}
		return result;
	}

	private Map<Integer, cn.excel.HandleField> handleFieldMap = new HashMap<>();

	/**
	 * 获取HandleField
	 * 
	 * @param excelField
	 * @return
	 */
	private cn.excel.HandleField getHandleField(ExcelField excelField) {
		cn.excel.HandleField handleField = handleFieldMap.get(excelField.sort());
		if (handleField == null) {
			try {
				Class<?> appContext = Class.forName("leap.core.AppContext");
				Method method = appContext.getMethod("getBean", Class.class);
				handleField = (cn.excel.HandleField) method.invoke(null, excelField.handleField());
			} catch (Exception ex) {
				try {
					handleField = (cn.excel.HandleField) excelField.handleField().newInstance();
				} catch (Exception e) {
					logger.warn("配置的HandleField类型错误");
				}
			}
			handleField.setParams(this.params);
			handleFieldMap.put(excelField.sort(), handleField);
		}
		return handleField;
	}

	private boolean checkHandle(Cell cell, Field field, ExcelField excelField, Map<String, Object> data) {
		// 调用处理器进行校验
		if (cn.excel.HandleField.class.isAssignableFrom(excelField.handleField())) {
			cn.excel.HandleField handleField = this.getHandleField(excelField);
			if (!handleField.validate(field.getName(), data)) {
				String msg = handleField.getMessage();
				this.flagError(cell, msg);
				return false;
			}
		}
		return true;
	}

	/**
	 * 检查格式
	 * 
	 * @param cell
	 * @param field
	 * @param value
	 * @return
	 */
	private boolean checkFormat(Cell cell, Field field, ExcelField excelField, String value) {
		String format = excelField.format();
		boolean result = true;
		String message = "格式不正确";

		if (StringUtils.isNotEmpty(format)) {
			result = true;
			if ("string".equalsIgnoreCase(format)) {
				result = true;
			} else if ("datetime".equalsIgnoreCase(format) && DateUtils.toTimestamp(value) == null) {
				result = false;
				message = "格式不正确，需精确到时分秒";
			} else if ("date".equalsIgnoreCase(format) && DateUtils.toDate(value) == null) {
				result = false;
				message = "格式不正确，需精确到年月日";
			} else if ("time".equalsIgnoreCase(format) && DateUtils.toTime(value) == null) {
				result = false;
				message = "格式不正确，请输入时间格式";
			} else if (("double".equalsIgnoreCase(format) || "float".equalsIgnoreCase(format)) && !NumberUtils.isNumber(value)) {
				result = false;
				message = "格式不正确，请输入数值";
			} else if ("money".equalsIgnoreCase(format)) {
				String strValue = StringUtils.removeEnd(value, "0");
				String[] values = StringUtils.split(strValue, ".");
				if (StringUtils.startsWith(value, "-") || !NumberUtils.isNumber(value) || (values.length == 2 && values[1].length() > 2)) {
					result = false;
					message = "格式不正确，请输入金额";
				}
			} else if (("byte".equalsIgnoreCase(format) || "short".equalsIgnoreCase(format) || "int".equalsIgnoreCase(format) || "long".equalsIgnoreCase(format))
					&& !NumberUtils.isDigits(value)) {
				result = false;
				message = "格式不正确，请输入整数";
			} else if ("mobile".equalsIgnoreCase(format) && !Pattern.compile("^1[3|4|5|7|8][0-9]{9}$").matcher(value).find()) {
				result = false;
				message = "格式不正确，请输入手机号";
			} else if ("tags".equals(format)) {
			    if (StringUtils.isNotEmpty(value)) {
                    String[] tags = {"排放点","污染点","扩散点"};
                    List list = Arrays.asList(tags);
                    String[] str = value.split(",");
                    w:for (String tag : str) {
                        if (!list.contains(tag)) {
                            result = false;
                            message = "格式不正确，请输入正确的分类格式";
                            break w;
                        }
                    }
                }
			} else if (StringUtils.startsWith(format, "{") && StringUtils.endsWith(format, "}")) {
				try {
					String json = StringUtils.replace(format, "\\", "\\\\");
					JSONObject obj = JSON.parseObject(json);
					String regex = obj.getString("regex");
					String msg = obj.getString("msg");
					if (StringUtils.isNotEmpty(regex)) {
						Pattern pattern = regexMap.get(field.getName());
						pattern = (pattern == null) ? Pattern.compile(regex) : pattern;
						regexMap.put(field.getName(), pattern);
						result = pattern.matcher(value).find();
						if (result == false) {
							this.flagError(cell, msg);
							return false;
						}
					}
				} catch (Exception ex) {
					logger.warn("配置JSON格式错误：{}", format);
				}
			}
			if (result == false) {
				this.flagError(cell, message);
			}
		}
		return result;
	}

	// 用于检查是否唯一
	private Map<String, List<String>> uniqueMap = new HashMap<>();
	private Map<String, Pattern> regexMap = new HashMap<>();

	/**
	 * 清理红色背景色
	 * 
	 * @param cell
	 */
	private void clearError(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		if (cellStyle != null) {
			CellStyle newStyle = this.wb.createCellStyle();
			newStyle.cloneStyleFrom(cellStyle);
			newStyle.setFillPattern(FillPatternType.NO_FILL);

			Font font = this.wb.getFontAt(newStyle.getFontIndex());

			Font newFont = this.wb.createFont();
			newFont.setColor(IndexedColors.BLACK.index);
			newFont.setFontHeightInPoints(font.getFontHeightInPoints());
			newFont.setFontName(font.getFontName());
			newFont.setCharSet(font.getCharSet());
			newStyle.setFont(newFont);

			cell.setCellStyle(newStyle);
		}
	}

	/**
	 * 标记错误(标红，另外加批注)
	 */
	private void flagError(Cell cell, String message) {
		// 标红
		CellStyle cellStyle = cell.getCellStyle();
		if (cellStyle != null) {
			CellStyle newStyle = this.wb.createCellStyle();
			newStyle.cloneStyleFrom(cellStyle);
			Font font = this.wb.getFontAt(newStyle.getFontIndex());

			Font newFont = this.wb.createFont();
			newFont.setColor(IndexedColors.AUTOMATIC.index);
			newFont.setFontHeightInPoints(font.getFontHeightInPoints());
			newFont.setFontName(font.getFontName());
			newFont.setCharSet(font.getCharSet());
			newFont.setBold(false);
			newStyle.setFont(newFont);
			newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			newStyle.setFillForegroundColor(IndexedColors.TAN.index);

			// newStyle.setFillForegroundColor(IndexedColors.RED.index);
			// newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell.setCellStyle(newStyle);
		}
		// 加错误批注
		this.addComment(cell, message);
	}

	private void addComment(Cell cell, String message) {
		int col1 = cell.getColumnIndex();
		int row1 = cell.getRowIndex();
		int col2 = col1 + 3;
		int row2 = row1 + 3;
		ClientAnchor clientAnchor = patriarch.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
		Comment comment = patriarch.createCellComment(clientAnchor);
		RichTextString richTextString = wb.getCreationHelper().createRichTextString(message);
		comment.setString(richTextString);
	}

	public <E> List<E> getModelList(Class<E> clazz) throws InstantiationException, IllegalAccessException, InvocationTargetException {
		List<E> entityList = new ArrayList<>();
		for (Map<String, Object> data : dataList) {
			if (data != null && !data.isEmpty()) {
				E entity = clazz.newInstance();
				for (Object[] objs : annotationList) {
					ExcelField excelField = (ExcelField) objs[0];
					Field field = (Field) objs[1];
					String fieldName = field.getName();
					Object fieldValue = data.get(fieldName);
					if (cn.excel.HandleField.class.isAssignableFrom(excelField.handleField())) {
						cn.excel.HandleField handleField = this.getHandleField(excelField);
						fieldValue = handleField.translate(data);
					}
					this.setFieldValue(entity, field, fieldValue);
				}
				entityList.add(entity);
			}
		}
		return entityList;
	}

	private void setFieldValue(Object obj, Field field, Object fieldValue) throws IllegalAccessException, InvocationTargetException {
		if (fieldValue != null) {
			Class<?> fieldType = field.getType();
			String strValue = String.valueOf(fieldValue);
			Object propertyValue = fieldValue;
			if (fieldType.isAssignableFrom(Date.class)) {
				propertyValue = DateUtils.toTimestamp(strValue);
			} else if (fieldType.isAssignableFrom(Double.class)) {
				propertyValue = new Double(strValue);
			} else if (fieldType.isAssignableFrom(Integer.class)) {
				propertyValue = new Integer(strValue);
			} else if (fieldType.isAssignableFrom(Boolean.class)) {
				propertyValue = new Boolean(strValue);
			} else if (fieldType.isAssignableFrom(java.sql.Time.class)) {
				propertyValue = DateUtils.toTime(strValue);
			} else if (fieldType.isAssignableFrom(java.sql.Date.class)) {
				propertyValue = DateUtils.toDate(strValue);
			}
			// 如果属性为空，则导入时忽略
			if (propertyValue != null) {
				try {
					BeanUtils.setProperty(obj, field.getName(), propertyValue);
				} catch (Exception ex) {
				}
			}
		}
	}

	/**
	 * 输出数据流
	 * 
	 * @param os
	 *            输出数据流
	 */
	public void write(OutputStream os) throws IOException {
		wb.write(os);
	}

	/**
	 * 输出到客户端
	 * 
	 * @param fileName
	 *            输出文件名
	 */
	public void write(HttpServletResponse response, String fileName) throws IOException {
		response.reset();
		response.setContentType("application/octet-stream; charset=utf-8");
		response.setHeader("Content-Disposition", "attachment; filename=\"" + java.net.URLEncoder.encode(fileName, "utf-8") + "\"");
		write(response.getOutputStream());
	}

	/**
	 * 输出到文件（文件必须为xlsx格式）
	 * 
	 * @param filePath
	 *            输出文件名
	 */
	public void write(String filePath) throws FileNotFoundException, IOException {
		FileOutputStream os = new FileOutputStream(filePath);
		this.write(os);
	}

	/**
	 * 获取行对象
	 * 
	 * @param rownum
	 * @return
	 */
	private Row getRow(int rownum) {
		return this.sheet.getRow(rownum);
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellValue(Cell cell) {
		CellType cellType = cell.getCellTypeEnum();
		if (cellType == CellType.STRING) {
			return cell.getStringCellValue();
		} else if (cellType == CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				short dataFormat = cell.getCellStyle().getDataFormat();
				if (dataFormat == 14) {
					return sdf14.format(date);
				} else if (dataFormat == 22) {
					return sdf22.format(date);
				} else if (dataFormat == 177) {
					return sdf177.format(date);
				}
				return new DataFormatter().formatCellValue(cell, evaluator);
			} else {
				double cellValue = cell.getNumericCellValue();
				String strValue = numberFormat.format(cellValue);
				if (StringUtils.endsWith(strValue, ".0")) {
					return StringUtils.removeEnd(strValue, ".0");
				} else {
					return strValue;
				}
			}
		} else if (cellType == CellType.BOOLEAN) {
			return cell.getBooleanCellValue() ? "True" : "False";
		} else if (cellType == CellType.FORMULA) {
			CellValue cellValue = evaluator.evaluate(cell);
			return this.getCellValue(cellValue);
		}
		return null;
	}

	/**
	 * 获取公式计算值
	 * 
	 * @param cellValue
	 * @return
	 */
	private String getCellValue(CellValue cellValue) {
		CellType cellType = cellValue.getCellTypeEnum();
		if (cellType == CellType.STRING) {
			return cellValue.getStringValue();
		} else if (cellType == CellType.NUMERIC) {
			double doubleValue = cellValue.getNumberValue();
			String strValue = numberFormat.format(doubleValue);
			if (StringUtils.endsWith(strValue, ".0")) {
				return StringUtils.removeEnd(strValue, ".0");
			} else {
				return strValue;
			}
		} else if (cellType == CellType.BOOLEAN) {
			return cellValue.getBooleanValue() ? "True" : "False";
		}
		return null;
	}

	public int getDataNum() {
		return dataNum;
	}

	public int getSuccessNum() {
		return successNum;
	}

	public int getErrorNum() {
		return errorNum;
	}

	public static void main(String[] args) {
		boolean result = NumberUtils.isDigits("H");
		System.out.println(result);
	}
}
