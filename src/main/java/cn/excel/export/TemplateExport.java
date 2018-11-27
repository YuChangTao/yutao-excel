package cn.excel.export;

import cn.excel.command.MergeCommand;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * jxls 模板导出
 *
 * <example>
 *
 *   String path = this.getClass().getClassLoader().getResource("template/" + fileName).toURI().getPath();
 *   TemplateExport templateExport = new TemplateExport(path, list);
 *
 * </example>
 * @author yutyi
 */
public class TemplateExport {

	private Object data;

	private PoiTransformer transformer;

    /**
     *
     * @param templateFile 模板文件路径
     * @param data
     * @throws InvalidFormatException
     * @throws IOException
     */
	public TemplateExport(String templateFile, Object data) throws InvalidFormatException, IOException {
		this(templateFile, data, false);
		
	}

    /**
     *
     * @param templateFile 模板文件路径
     * @param data
     * @param isSXSSF
     * @throws IOException
     * @throws InvalidFormatException
     */
	public TemplateExport(String templateFile, Object data, boolean isSXSSF) throws IOException, InvalidFormatException {
		this.data = data;
		Workbook workbook = WorkbookFactory.create(new File(templateFile));
		XlsCommentAreaBuilder.addCommandMapping("merge", MergeCommand.class);
		
		if (isSXSSF) {
			transformer = PoiTransformer.createSxssfTransformer(workbook);
		} else {
			transformer = PoiTransformer.createTransformer(workbook);
		}
	}

    /**
     * 支持jar包运行的项目（无法获取jar包中文件的路径，因为jar本身就是一个文件，而非目录）
     *
     * @param inputStream 模板文件流
     * @param data
     * @param isSXSSF
     * @throws IOException
     * @throws InvalidFormatException
     */
    public TemplateExport(InputStream inputStream, Object data, boolean isSXSSF) throws IOException, InvalidFormatException {
        this.data = data;
        Workbook workbook = WorkbookFactory.create(inputStream);
        //添加自定义合并指令
        XlsCommentAreaBuilder.addCommandMapping("merge", MergeCommand.class);

        if (isSXSSF) {
            transformer = PoiTransformer.createSxssfTransformer(workbook);
        } else {
            transformer = PoiTransformer.createTransformer(workbook);
        }
    }

	public Object getData() {
		return data;
	}

	public void setData(Object data) {
		this.data = data;
	}

	public Transformer getTransformer() {
		return transformer;
	}

	/**
	 * 输出数据流
	 * 
	 * @param os
	 *            输出数据流
	 */
	public TemplateExport write(OutputStream os) throws IOException {
		transformer.setOutputStream(os);
		Context context = new Context();
		context.putVar("data", this.data);
		JxlsHelper.getInstance().processTemplate(context, transformer);
		return this;
	}

	/**
	 * 输出到客户端
	 * 
	 * @param fileName
	 *            输出文件名
	 */
	public TemplateExport write(HttpServletResponse response, String fileName) throws IOException {
		response.reset();
		response.setContentType("application/octet-stream; charset=utf-8");
		response.setHeader("Content-Disposition", "attachment; filename=\"" + java.net.URLEncoder.encode(fileName, "utf-8") + "\"");
		write(response.getOutputStream());
		return this;
	}

	/**
	 * 输出到文件（文件必须为xlsx格式）
	 * 
	 * @param filePath
	 *            输出文件名
	 */
	public TemplateExport write(String filePath) throws IOException {
		FileOutputStream os = new FileOutputStream(filePath);
		this.write(os);
		return this;
	}
}
