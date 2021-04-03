package org.baolin.treeexcel;

import org.apache.poi.ss.usermodel.*;

import java.util.Map;

/**
 * @author 钟宝林
 **/
final class ExcelStyle {

    /**
     * 表名样式
     */
    static CellStyle tableNameStyle(Workbook workbook, Map<String, CellStyle> styleMap) {
        String key = "tableNameStyle";
        CellStyle style = styleMap.get(key);
        if (style == null) {
            style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(true);
            Font font = workbook.createFont();
            font.setFontName("黑体");
            font.setFontHeightInPoints((short) 20);
            style.setFont(font);
            styleMap.put(key, style);
        }
        return style;
    }

    /**
     * 表头样式
     */
    static CellStyle headerStyle(Workbook workbook, Map<String, CellStyle> styleMap) {
        String key = "headerStyle";
        CellStyle style = styleMap.get(key);
        if (style == null) {
            style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(true);
            //设置边框
            style.setBorderBottom(BorderStyle.THIN); // 底部边框
            style.setBorderLeft(BorderStyle.THIN);  // 左边边框
            style.setBorderRight(BorderStyle.THIN); // 右边边框
            style.setBorderTop(BorderStyle.THIN); // 上边边框
            Font font = workbook.createFont();
            font.setFontName("黑体");
            style.setFont(font);
            styleMap.put(key, style);
        }
        return style;
    }

    /**
     * 内容样式
     */
    static CellStyle contentStyle(Workbook workbook, Map<String, CellStyle> styleMap) {
        String key = "contentStyle";
        CellStyle style = styleMap.get(key);
        if (style == null) {
            style = workbook.createCellStyle();
            style.setWrapText(true);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            //设置边框
            style.setBorderBottom(BorderStyle.THIN); // 底部边框
            style.setBorderLeft(BorderStyle.THIN);  // 左边边框
            style.setBorderRight(BorderStyle.THIN); // 右边边框
            style.setBorderTop(BorderStyle.THIN); // 上边边框
            styleMap.put(key, style);
        }
        return style;
    }

    /**
     * 数字内容样式
     */
    static CellStyle numberStyle(Workbook workbook, Map<String, CellStyle> styleMap) {
        String key = "numberStyle";
        CellStyle style = styleMap.get(key);
        if (style == null) {
            style = workbook.createCellStyle();
            style.setWrapText(true);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            //设置边框
            style.setBorderBottom(BorderStyle.THIN); // 底部边框
            style.setBorderLeft(BorderStyle.THIN);  // 左边边框
            style.setBorderRight(BorderStyle.THIN); // 右边边框
            style.setBorderTop(BorderStyle.THIN); // 上边边框
//        DataFormat df = workbook.createDataFormat();  //此处设置数据格式
//        style.setDataFormat(df.getFormat("#,#0.00"));; //数据格式只显示整数，如果是小数点后保留两位，可以写contentStyle.setDataFormat(df.getFormat("#,#0.00"));
            styleMap.put(key, style);
        }
        return style;
    }

    /**
     * 合计样式
     */
    static CellStyle totalStyle(Workbook workbook, Map<String, CellStyle> styleMap) {
        String key = "totalStyle";
        CellStyle style = styleMap.get(key);
        if (style == null) {
            style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontName("黑体");
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setWrapText(true);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            //设置边框
            style.setBorderBottom(BorderStyle.THIN); // 底部边框
            style.setBorderLeft(BorderStyle.THIN);  // 左边边框
            style.setBorderRight(BorderStyle.THIN); // 右边边框
            style.setBorderTop(BorderStyle.THIN); // 上边边框
            styleMap.put(key, style);
        }
        return style;
    }

    /**
     * 合计样式
     */
    static CellStyle totalNumberStyle(Workbook workbook, Map<String, CellStyle> styleMap) {
        String key = "totalNumberStyle";
        CellStyle style = styleMap.get(key);
        if (style == null) {
            style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontName("黑体");
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setWrapText(true);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            //设置边框
            style.setBorderBottom(BorderStyle.THIN); // 底部边框
            style.setBorderLeft(BorderStyle.THIN);  // 左边边框
            style.setBorderRight(BorderStyle.THIN); // 右边边框
            style.setBorderTop(BorderStyle.THIN); // 上边边框
            styleMap.put(key, style);
        }
        return style;
    }

}
