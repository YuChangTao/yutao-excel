package cn.excel.export;

import cn.excel.ExcelField;
import org.apache.commons.lang.ArrayUtils;
import org.apache.commons.lang.ObjectUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * MapList/ModelList导出成为Excel 2017版本，需要指定后缀为xlsx <br/>
 * ExcelExport export = new ExcelExport(columnNames, keys, dataList); <br/>
 * export.write(respose, "test.xlsx");
 * 
 * @author yutyi
 *
 */
public class ExcelExport {

	/**
	 * 工作薄对象
	 */
	private SXSSFWorkbook wb;

	/**
	 * 工作表对象
	 */
	private Sheet sheet;

	/**
	 * 样式列表
	 */
	private Map<String, CellStyle> styles;

	/**
	 * 当前行号(0-based)
	 */
	private int rownum;

	/**
	 * 导出excel表头集合
	 */
	private String[] columns;

	/**
	 * 数据列Map中key集合
	 */
	private String[] keys;

	/**
	 * 数据集合
	 */
	private List<Map<String, Object>> dataList;

    /**
     * 注解@ExcelField属性集合
     */
	private List<Object[]> annotationList;

	/**
	 * 自定义表头名（使用，分割）和key导出
     *
	 * @param columns 导出excel表头集合
	 * @param keys 数据列Map中key集合
	 * @param dataList 数据集合
	 */
	public ExcelExport(String columns, String keys, List<Map<String, Object>> dataList) {
		String[] columnNames = StringUtils.split(columns, ",");
		String[] keyNames = StringUtils.split(keys, ",");
		if (columnNames.length != keyNames.length) {
			throw new RuntimeException("常规导出时，数据列出Map中的Key数量不一致");
		}
		this.columns = columnNames;
		this.keys = keyNames;
		this.dataList = dataList;

		this.initialize();
	}

    /**
     * 通过类注解@ExcelField导出
     *
     * @param clazz
     * @param dataList
     */
	public ExcelExport(Class clazz,List<Map<String,Object>> dataList) {
        this.dataList = dataList;
        this.initialize(clazz);
    }

    /**
     * 初始化Excel
     */
	private void initialize() {
		this.wb = new SXSSFWorkbook(500);
		this.sheet = wb.createSheet("Sheet1");
		this.styles = createStyles(wb);
		// 创建第一行空白行
		Row firstRow = sheet.createRow(rownum++);
		firstRow.setHeightInPoints(14);

		// 创建列头
		new ExcelHeader(this.columns).create();
		sheet.setColumnWidth(0, 2 * 256);

		// 循环写入数据
		for (Map<String, Object> data : dataList) {
			Row dataRow = sheet.createRow(rownum++);
			this.addBlankCell(dataRow);
			int cellIndex = 1;
			for (String key : this.keys) {
				Object value = data.get(key);
				this.addCell(dataRow, cellIndex++, value);
			}
		}
	}

    /**
     * 注解类初始化Excel
     * @param clazz
     */
    private void initialize(Class clazz) {
        this.wb = new SXSSFWorkbook(500);
        this.sheet = wb.createSheet("Sheet1");
        this.styles = createStyles(wb);
        annotationList = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();

        //添加注解@ExcelField的属性
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null) {
                annotationList.add(new Object[]{field,excelField});
            }
        }

        //通过注解排序
        if (annotationList !=null && annotationList.size()>0) {
            annotationList.sort((o1, o2) -> ((ExcelField) o1[1]).sort() < ((ExcelField) o2[1]).sort() ? 1:0);
        }

        //表头集合
        List<String> headList = new ArrayList<>();
        for (Object[] obj : annotationList) {
            ExcelField excelField = (ExcelField) obj[1];
            headList.add(excelField.columnName());
        }

        createHead(headList);
        setDataList();
    }

    /**
     * 创建表头
     * @param headList
     */
    private void createHead(List<String> headList) {
        if (headList != null && headList.size() > 0 ) {
            Row row = sheet.createRow(rownum++);
            for (int column = 0; column < headList.size(); column++) {
                //添加单元格
                addCell(row,column,headList.get(column));
                row.getCell(column).setCellStyle(styles.get("header"));
            }
        }
    }

    /**
     * 设置数据
     */
    public void setDataList() {
        if (dataList != null && dataList.size() > 0 ) {
            for (int column = 0; column < dataList.size(); column++) {
                //填充数据
                Row row = this.sheet.createRow(rownum++);
                for (int i = 0; i < annotationList.size(); i++) {
                    Field field = (Field)annotationList.get(i)[0];
                    addCell(row,column,dataList.get(column).get(field.getName()));
                }
            }
        }
    }

	/**
	 * 添加一行
	 * 
	 * @return 行对象
	 */
	public Row addRow() {
		return sheet.createRow(rownum++);
	}

	/**
	 * 添加空白单元格
	 * 
	 * @param row
	 */
	public void addBlankCell(Row row) {
		row.createCell(0).setCellStyle(styles.get("blank"));
	}
	
	/**
	 * 设置列宽
	 * @param columnIndex
	 * @param width
	 */
	public void setColumnWidth(int columnIndex, double width){
		int finalWidth = new Double(width * 256 + 184).intValue();
		sheet.setColumnWidth(columnIndex, finalWidth);
	}

	/**
	 * 添加单元格并填充数据
	 * 
	 * @param row
	 * @param cellIndex
	 * @param cellValue
	 */
	public void addCell(Row row, int cellIndex, Object cellValue) {
		Cell cell = row.createCell(cellIndex);
		cell.setCellStyle(styles.get("data_auto"));
		if (cellValue != null) {
			if (cellValue instanceof java.sql.Date) {
				cell.setCellValue((Date) cellValue);
				cell.setCellStyle(styles.get("data_date"));
				this.setColumnWidth(cellIndex, 9);
			} else if (cellValue instanceof java.sql.Time) {
				cell.setCellValue((Date) cellValue);
				cell.setCellStyle(styles.get("data_time"));
				this.setColumnWidth(cellIndex, 9);
			} else if (cellValue instanceof java.util.Date) {
				cell.setCellValue((Date) cellValue);
				cell.setCellStyle(styles.get("data_datetime"));
				this.setColumnWidth(cellIndex, 16);
			} else if (cellValue instanceof Calendar) {
				cell.setCellValue((Calendar) cellValue);
			} else if (cellValue instanceof Boolean) {
				cell.setCellValue((Boolean) cellValue);
			} else if (cellValue instanceof Short) {
                cell.setCellStyle(styles.get("data_number"));
				cell.setCellValue((Short) cellValue);
			} else if (cellValue instanceof Integer) {
                cell.setCellStyle(styles.get("data_number"));
			    cell.setCellValue((Integer) cellValue);
			} else if (cellValue instanceof Long) {
				cell.setCellStyle(styles.get("data_number"));
				cell.setCellValue((Long) cellValue);
			} else if (cellValue instanceof Float) {
                cell.setCellStyle(styles.get("data_decimal"));
				cell.setCellValue((Float) cellValue);
			} else if (cellValue instanceof Double) {
                cell.setCellStyle(styles.get("data_decimal"));
				cell.setCellValue((Double) cellValue);
			} else {
				String strValue = ObjectUtils.toString(cellValue);
				cell.setCellValue(strValue);
				if(strValue.length() > 20){
					cell.setCellStyle(styles.get("data_string"));
				}
			}
		}
	}

	/**
	 * 添加合并单元格
	 * 
	 * @param cellRangeAddress 合并单元格坐标
	 */
	public void addMergedRegion(CellRangeAddress cellRangeAddress) {
		sheet.addMergedRegion(cellRangeAddress);
	}

	/**
	 * 创建表格样式
	 * 
	 * @param wb
	 *            工作薄对象
	 * @return 样式列表
	 */
	private Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>(10);

		//创建单元格样式
		CellStyle style = wb.createCellStyle();
        //设置单元格字体样式
		Font blankFont = wb.createFont();
		blankFont.setFontHeightInPoints((short) 16);
		style.setFont(blankFont);
		styles.put("blank", style);

		//创建标题样式
		style = wb.createCellStyle();
		//设置水平居中
		style.setAlignment(HorizontalAlignment.CENTER);
		//设置垂直居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		//设置边框样式
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setWrapText(false);
		Font titleFont = wb.createFont();
		titleFont.setFontName("Courier New");
		titleFont.setFontHeightInPoints((short) 11);
		titleFont.setBold(true);
		style.setFont(titleFont);
		styles.put("title", style);

		//创建表头样式
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("title"));
		Font headerFont = wb.createFont();
		headerFont.setFontName("Courier New");
		headerFont.setFontHeightInPoints((short) 10);
		headerFont.setBold(true);
		style.setFont(headerFont);
		styles.put("header", style);

		//创建自动数据格式样式
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("title"));
		style.setAlignment(HorizontalAlignment.GENERAL);
		Font dataFont = wb.createFont();
		titleFont.setFontName("Courier New");
		dataFont.setFontHeightInPoints((short) 10);
		style.setFont(dataFont);
		styles.put("data_auto", style);

		//文本
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data_auto"));
		style.setWrapText(true);
		styles.put("data_string", style);

		//整数
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data_auto"));
		style.setDataFormat(wb.createDataFormat().getFormat("0"));
		styles.put("data_number",style);

		//小数
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data_auto"));
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("data_decimal",style);

		//java.sql.Date
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data_auto"));
		style.setDataFormat(wb.createDataFormat().getFormat("yyyy/MM/dd"));
		styles.put("data_date", style);

		//java.util.Date
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data_auto"));
		style.setDataFormat(wb.createDataFormat().getFormat("yyyy/MM/dd HH:mm:ss"));
		styles.put("data_datetime", style);

		//java.sql.Time
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data_auto"));
		style.setDataFormat(wb.createDataFormat().getFormat("HH:mm:ss"));
		styles.put("data_time", style);
		
		return styles;
	}

	/**
     * 输出数据流
     *
     * @param os
     *            输出数据流
     */
    public ExcelExport write(OutputStream os) throws IOException {
        wb.write(os);
        wb.dispose();
        return this;
    }

    /**
     * 输出到客户端
     *
     * @param fileName
     *            输出文件名
     */
    public ExcelExport write(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename=\"" + java.net.URLEncoder.encode(fileName, "utf-8") + "\"");
        this.write(response.getOutputStream());
        return this;
    }

    /**
     * 输出到文件（文件必须为xlsx格式）
     *
     * @param filePath
     *            输出文件名
     */
    public ExcelExport write(String filePath) throws IOException {
        FileOutputStream os = new FileOutputStream(filePath);
        this.write(os);
        return this;
    }

	/**
	 * Excel标题分析类，支持水平和垂直合并
	 * 
	 * @author zhengwenquan
	 *
	 */
	class ExcelHeader {
		private List<String> headerList;
		private Map<Integer, Row> rows = new HashMap<>(15);
		private List<Map<Integer, String>> multiRows = new ArrayList<>(10);
		private List<List<String>> multiColumns = new ArrayList<>(5);
		private Map<String, Integer[]> merges = new HashMap<>(5);
		// private Map<Integer, Boolean> comments = new HashMap<>();

		public ExcelHeader(List<String> headerList) {
			this.headerList = headerList;
		}

		public ExcelHeader(String[] headers) {
			this.headerList = Arrays.asList(headers);
		}

		public void create() {
		    //分析
			this.analyze();
			//创建单元格
			this.createCell();
			//水平合并
			this.horizontalMarge();
			//垂直合并
			this.verticalMarge();

			this.mergeRegion();
		}

		private void analyze() {
			for (int column = 1; column <= headerList.size(); column++) {
				String[] ss = StringUtils.split(headerList.get(column - 1), "#");
				// comments.put(column, ss.length == 1); // 首行显示批注
				ArrayUtils.reverse(ss);
				multiColumns.add(new ArrayList<>(Arrays.asList(ss)));
				for (int row = 0; row < ss.length; row++) {
					String text = ss[row];
					Map<Integer, String> xCells = multiRows.size() <= row ? new HashMap<>() : multiRows.get(row);
					xCells.put(column, text);
					if (multiRows.size() <= row) {
						multiRows.add(xCells);
					}
				}
			}

			// 向空白区域设置值
			int rowCount = multiRows.size();
			for (int column = 1; column <= multiColumns.size(); column++) {
				List<String> columns = multiColumns.get(column - 1);
				if (columns.size() < rowCount) {
					String lastValue = columns.get(columns.size() - 1);
					for (int i = columns.size(); i < rowCount; i++) {
						multiRows.get(i).put(column, lastValue);
						columns.add(lastValue);
					}
				}
				Collections.reverse(columns);
			}
			// 反转集合
			Collections.reverse(multiRows);
		}

		/**
		 * 填充单元格内容
		 */
		private void createCell() {
			for (int i = 0; i < multiRows.size(); i++) {
				Map<Integer, String> xCells = multiRows.get(i);
				Row row = addRow();
				rows.put(i, row);
				addBlankCell(row);
				for (Map.Entry<Integer, String> entry : xCells.entrySet()) {
					int column = entry.getKey();
					Cell cell = row.createCell(column);
					cell.setCellValue(entry.getValue());
					cell.setCellStyle(styles.get("header"));
				}
			}
		}

		/**
		 * 水平合并分析
		 */
		private void horizontalMarge() {
			int rowCount = multiRows.size();
			int columnCount = multiColumns.size();
			int startRow = rownum - rowCount;
			for (int row = rowCount - 1; row >= 0; row--) {
				int firstRow = startRow + row;
				int lastRow = startRow + row;
				int firstCol = 1;
				int lastCol = 1;
				String value = null;
				for (int column = 1; column <= columnCount; column++) {
					List<String> columns = multiColumns.get(column - 1);
					if (column == 1) {
						value = columns.get(row);
						continue;
					}
					if (columns.get(row).equals(value)) {
						lastCol = column;
						multiRows.get(row).put(column, "*");
					} else {
						value = columns.get(row);
						if (firstCol != lastCol) {
							Integer[] addr = merges.get((firstRow + 1) + ":" + firstCol);
							if (addr != null && addr[3] == lastCol) {
								lastRow = addr[1];
								merges.remove((firstRow + 1) + ":" + firstCol);
							}
							merges.put(firstRow + ":" + firstCol, new Integer[] { firstRow, lastRow, firstCol, lastCol });
						}
						firstCol = column;
						lastCol = firstCol;
						lastRow = firstRow;
					}
					if (column == columnCount && firstCol != lastCol) {
						Integer[] addr = merges.get((firstRow + 1) + ":" + firstCol);
						if (addr != null && addr[3] == lastCol) {
							lastRow = addr[1];
							merges.remove((firstRow + 1) + ":" + firstCol);
						}
						merges.put(firstRow + ":" + firstCol, new Integer[] { firstRow, lastRow, firstCol, lastCol });
					}
				}
			}
		}

		/**
		 * 垂直合并分析
		 */
		private void verticalMarge() {
			int rowCount = multiRows.size();
			int columnCount = multiColumns.size();
			int startRow = rownum - rowCount;
			for (int column = 1; column <= columnCount; column++) {
				int firstCol = column;
				int lastCol = column;
				int firstRow = startRow;
				int lastRow = startRow;
				String value = null;
				for (int row = 0; row < rowCount; row++) {
					Map<Integer, String> rows = multiRows.get(row);
					if (row == 0) {
						value = rows.get(column);
						continue;
					}
					if (!"*".equals(value) && rows.get(column).equals(value)) {
						lastRow = startRow + row;
					} else {
						if (!value.equals("*") && firstRow != lastRow) {
							Integer[] addr = merges.get(firstRow + ":" + firstCol);
							if (addr == null || lastRow != addr[1]) {
								addr = new Integer[] { firstRow, lastRow, firstCol, lastCol };
							}
							merges.put(firstRow + ":" + firstCol, addr);
						}
						value = rows.get(column);
						firstRow = startRow + row;
						lastRow = firstRow;
					}
					if (row == rowCount - 1 && firstRow != lastRow) {
						Integer[] addr = merges.get(firstRow + ":" + firstCol);
						if (addr == null || lastRow != addr[1]) {
							addr = new Integer[] { firstRow, lastRow, firstCol, lastCol };
						}
						merges.put(firstRow + ":" + firstCol, addr);
					}
				}
			}
		}

		private void mergeRegion() {
			for (Integer[] addr : merges.values()) {
				addMergedRegion(new CellRangeAddress(addr[0], addr[1], addr[2], addr[3]));
			}
		}
	}
}
