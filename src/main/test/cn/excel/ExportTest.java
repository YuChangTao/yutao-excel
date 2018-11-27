package cn.excel;

import cn.excel.export.ExcelExport;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 导出测试
 *
 * @author yutyi
 * @date 2018/11/07
 */
public class ExportTest {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String columns = "类型id,类型名称,状态,创建人,创建时间";
        String keys = "typeId,typeName,state,createId,createTime";
        List<Map<String,Object>> list = new ArrayList<>();
        Map<String,Object> map = new HashMap<>();
        map.put("typeId",1539730672078001L);
        map.put("typeName","粉尘检测仪");
        map.put("state",1);
        map.put("createId",1);
        map.put("createTime","2018-10-17");
        list.add(map);

        ExcelExport excelExport = new ExcelExport(columns,keys,list);
        FileOutputStream outputStream = new FileOutputStream(new File("G://2.xlsx"));
        excelExport.write(outputStream);
//        TemplateExport templateExport = new TemplateExport("G://1.xlsx",list,false);
//        templateExport.write(outputStream);
    }
}
