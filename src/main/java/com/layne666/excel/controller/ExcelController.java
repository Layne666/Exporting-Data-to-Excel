package com.layne666.excel.controller;

import com.layne666.excel.bean.Excel;
import com.layne666.excel.bean.User;
import com.layne666.excel.util.ExportExcelUtil;
import com.layne666.excel.util.ExportUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author layne666
 */
@Controller
public class ExcelController {
    @RequestMapping("/exportExcel")
    public void exportExcel(HttpServletResponse resp) {
        String[] colTitles = {"姓名","年龄","地址"};
        String[] properties ={"name","age","address"};

        //User对象和Person对象，属性仅有部分相同，若拿相同属性去导出，则可以导出
        User u1 = new User("张三",20,"北京",175);
        User u2 = new User("李四",22,"上海",178);
        User u3 = new User("王五",24,"广州",180);
        List<Object> list = new ArrayList<>();
        list.add(u1);list.add(u2);list.add(u3);

        Map<String, Object> result = ExportExcelUtil.exportExcel(list, colTitles, properties, "用户列表", "用户数据统计表", resp);
        System.out.println(result.get("msg"));
        System.out.println(result.get("success"));
    }

    @RequestMapping("/export")
    public void export(HttpServletResponse resp){
        Excel excel = new Excel();
        excel.setFileName("文件");
        excel.setHeadTitle("衣服表");
        Map<Integer, String[]> map = new HashMap<>(2);
        String[] str1 = {"序号","件数","111","合计"};
        map.put(1,str1);
        String[] str2 = {null,"上衣","裤子",null};
        map.put(2,str2);
        excel.setColTitles(map);
        excel.addHbdyg(1,2,0,0);
        excel.addHbdyg(1,1,1,2);
        excel.addHbdyg(1,2,3,3);
        ExportUtil.exportExcel(excel,resp);
    }
}
