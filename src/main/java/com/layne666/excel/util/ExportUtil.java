package com.layne666.excel.util;

import com.layne666.excel.bean.Excel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExportUtil {
        public static void exportExcel(Excel excel, HttpServletResponse resp) {
            //1.创建一个excel文件
            HSSFWorkbook workbook = new HSSFWorkbook();
            //1.2 创建头标题样式
            HSSFCellStyle headStyle = createCellStyle(workbook,(short)16,true);
            //1.3 创建列标题样式
            HSSFCellStyle colStyle = createCellStyle(workbook,(short)13,true);
            //1.4 创建正文样式
            HSSFCellStyle valueStyle = createCellStyle(workbook,(short)12,false);
            //自动换行
            valueStyle.setWrapText(true);

            //2.创建工作簿
            HSSFSheet sheet = workbook.createSheet();

            CellRangeAddress head = new CellRangeAddress(0,0,0,excel.getColTitles().get(1).length-1);
            sheet.addMergedRegion(head);

            //2.1 加载合并单元格对象
            List<CellRangeAddress> hbdygList = excel.getHbdygList();
            if(hbdygList!=null){
                for (CellRangeAddress cas : hbdygList) {
                    sheet.addMergedRegion(cas);
                }
            }
            //2.2 设置默认列宽
            sheet.setDefaultColumnWidth(25);

            //3.创建行
            //3.1 创建头标题行;并且设置头标题
            HSSFRow row1 = sheet.createRow(0);
            //3.2 创建单元格
            HSSFCell cell = row1.createCell(0);
            //3.3 设置单元格样式
            cell.setCellStyle(headStyle);
            //3.4 设置单元格值
            cell.setCellValue(excel.getHeadTitle());

            //4.创建行
            //4.1 创建列标题行;并且设置列标题
            Map<Integer, String[]> colTitles = excel.getColTitles();
            for (Map.Entry<Integer, String[]> entry : colTitles.entrySet()) {
                HSSFRow row = sheet.createRow(entry.getKey());
                for (int i = 0;i < 4; i++){
                    //4.2 创建单元格
                    HSSFCell cell2 = row.createCell(i);
                    //4.3 设置单元格样式
                    cell2.setCellStyle(colStyle);
                    //4.4 设置单元格值
                    cell2.setCellValue(entry.getValue()[i]);
                }
            }


            //5.获取输出流对象
            ServletOutputStream out = null;
            try {
                out = resp.getOutputStream();
                //6.添加时间，防止文件名字重复
                SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timer = sdf.format(new Date(System.currentTimeMillis()));
                //7.设置信息头
                resp.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(excel.getFileName()+"("+timer+")","UTF-8")+".xls");
                resp.setHeader("Content-Type", "application/octet-stream");
                //8.写出文件
                workbook.write(out);
                //9.返回信息
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if(out != null){
                    try {
                        //10.关闭输出流
                        out.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

        }

        /**
         * 设置字体样式
         * @param workbook  excel文件
         * @param fontsize  字体大小
         * @param isBold    字体是否需要加粗
         * @return 样式
         */
        private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook, short fontsize, Boolean isBold) {
            HSSFCellStyle style = workbook.createCellStyle();
            //水平居中
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            //垂直居中
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            //创建字体
            HSSFFont font = workbook.createFont();
            //字体是否需要加粗
            if(isBold){
                //加粗
                font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            }
            //设置字体大小
            font.setFontHeightInPoints(fontsize);
            //加载字体
            style.setFont(font);
            return style;
        }

        /**
         *  通过反射获取对象的getter方法，来获取对象的属性值
         * @param obj   对象
         * @param name  对象的属性名称
         * @return      返回对象的属性值
         */
        private static Object getGetMethod(Object obj, String name){
            Method[] m = obj.getClass().getMethods();
            for(int i = 0;i < m.length;i++){
                if(("get"+name).toLowerCase().equals(m[i].getName().toLowerCase())){
                    try {
                        return m[i].invoke(obj);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
            return null;
        }
}
