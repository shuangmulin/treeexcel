package org.baolin.treeexcel;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @author 钟宝林
 * @since 2021/3/28/028 9:44
 **/
public class TreeExcel {

    public static Workbook getExcel(Table table) {
        return ExcelUtils.toExcel(table);
    }

    public static <T> Workbook getExcel(Class<T> clazz, List<T> dataList) {
        return null;
    }

}
