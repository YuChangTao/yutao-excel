package cn.excel.demo;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.excel.export.TemplateExport;
import org.apache.commons.lang.math.RandomUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;


public class Main {

	public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {

		List<Map<String, Object>> list = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			int num = RandomUtils.nextInt(5);
			num ++;
			
			Map<String, Object> map = new HashMap<>();
			map.put("name", "dept" + i);
			map.put("index", i + 1);
			map.put("num", num);
			List<Map<String, Object>> userList = new ArrayList<>();
			for (int j = 0; j < num; j++) {
				Map<String, Object> map2 = new HashMap<>();
				map2.put("name", "name_" + i + "_" + j);
				map2.put("age", RandomUtils.nextInt(90));
				userList.add(map2);
			}
			if(userList.size() > 0){
				userList.get(0).put("num", num);
			}
			map.put("userList", userList);
			list.add(map);
		}
		
		Map<String, Object> map = new HashMap<>();
		map.put("deptList", list);
		
		String path = "G:\\1.xlsx";
		TemplateExport templateExport = new TemplateExport(path, map);
		templateExport.write("G:\\1.xlsx");
		//templateExport.write(response().getServletResponse(), DateUtils.format("yyyyMMddHHmm") + "-派车统计表.xlsx");



	}

}
