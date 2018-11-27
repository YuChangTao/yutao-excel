package cn.excel.command;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiCellData;
import org.jxls.transform.poi.PoiTransformer;

/**
 * 合并指令
 * @author yutyi
 */
public class MergeCommand extends AbstractCommand {
	
	private String cols;
	private String rows;
	private Area area;
	private CellStyle cellStyle;

	@Override
	public String getName() {
		return "merge";
	}
	
	@Override
	public Command addArea(Area area) {
        if (super.getAreaList().size() >= 1) {
            throw new IllegalArgumentException("You can add only a single area to 'merge' command");
        }
        this.area = area;
        return super.addArea(area);
	}

	@Override
	public Size applyAt(CellRef cellRef, Context context) {
		int rows = 1, cols = 1;
		if(StringUtils.isNotBlank(this.rows)){
            Object rowsObj = getTransformationConfig().getExpressionEvaluator().evaluate(this.rows, context.toMap());
            if(rowsObj != null && NumberUtils.isDigits(rowsObj.toString())){
                rows = NumberUtils.toInt(rowsObj.toString());
            }
        }
        if(StringUtils.isNotBlank(this.cols)){
            Object colsObj = getTransformationConfig().getExpressionEvaluator().evaluate(this.cols, context.toMap());
            if(colsObj != null && NumberUtils.isDigits(colsObj.toString())){
                cols = NumberUtils.toInt(colsObj.toString());
            }
        }
        if(rows > 1 || cols > 1){
        	Transformer transformer = this.getTransformer();
        	if(transformer instanceof PoiTransformer){
        		return poiMerge(cellRef, context, (PoiTransformer)transformer, rows, cols);
        	}
        }
        
        area.applyAt(cellRef, context);
		
		return new Size(1, 1);
	}
	
	
	protected Size poiMerge(CellRef cellRef, Context context, PoiTransformer transformer, int rows, int cols) {
		Sheet sheet = transformer.getWorkbook().getSheet(cellRef.getSheetName());
		int firstRow = cellRef.getRow();
		int firstCol = cellRef.getCol();
		CellRangeAddress region = new CellRangeAddress(firstRow, firstRow + rows - 1, firstCol, firstCol + cols - 1);
		sheet.addMergedRegion(region);
		
		area.applyAt(cellRef, context);
		
		if(cellStyle == null){
			PoiCellData cellData = (PoiCellData)transformer.getCellData(cellRef);
			if(cellData != null){
				cellStyle = cellData.getCellStyle();
			}else{
				cellStyle = sheet.getWorkbook().createCellStyle();
				cellStyle.setBorderTop(BorderStyle.THIN);
				cellStyle.setBorderRight(BorderStyle.THIN);
				cellStyle.setBorderBottom(BorderStyle.THIN);
				cellStyle.setBorderLeft(BorderStyle.THIN);
				cellStyle.setTopBorderColor(IndexedColors.BLACK.index);
				cellStyle.setRightBorderColor(IndexedColors.BLACK.index);
				cellStyle.setBottomBorderColor(IndexedColors.BLACK.index);
				cellStyle.setLeftBorderColor(IndexedColors.BLACK.index);
				cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			}
		}
		//设置单元格样式
		for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
			Row row = sheet.getRow(i);
			if(row == null){
				row = sheet.createRow(i);
			}
			for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
				Cell cell = row.getCell(j);
				if(cell == null){
					cell = row.createCell(j);
				}
				if(cellStyle != null){
					cell.setCellStyle(cellStyle);
				}
			}
		}
		return new Size(cols, rows);
	}
	

	public String getCols() {
		return cols;
	}

	public void setCols(String cols) {
		this.cols = cols;
	}

	public String getRows() {
		return rows;
	}

	public void setRows(String rows) {
		this.rows = rows;
	}

	public Area getArea() {
		return area;
	}

	public void setArea(Area area) {
		this.area = area;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	
}
