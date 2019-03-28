package com.layne666.excel.bean;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author Layne666
 */
@Setter
@Getter
public class Excel implements Serializable {
    /**
     * 导出时的文件名称
     */
    public String fileName;
    /**
     * 表格中间头标题
     */
    private String headTitle;

    /**
     * 单元格的列标题
     * key 第几行
     * value 每行的标题数组
     */
    private Map<Integer,String[]> colTitles;

    /**
     * 合并单元格的列标题
     */
    private List<CellRangeAddress> hbdygList = new ArrayList<>();

    /**
     * 对象的属性
     */
    private String[] properties;

    /**
     * 需要导出的对象集合
     */
    private List<Object> objs;

    /**
     * 添加合并单元格对象
     */
    public void addHbdyg(int firstRow, int lastRow, int firstCol, int lastCol){
        CellRangeAddress cra = new CellRangeAddress(firstRow,lastRow,firstCol,lastCol);
        this.hbdygList.add(cra);
    }
}
