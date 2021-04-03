package org.baolin.treeexcel;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.baolin.treeexcel.utils.HeaderUtils;
import org.baolin.treeexcel.utils.NumberUtils;
import org.baolin.treeexcel.utils.StringUtils;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel工具类
 *
 * @author 钟宝林
 **/
@SuppressWarnings("rawtypes")
public final class ExcelUtils {
    /**
     * 默认行高
     */
    private static final short ROW_HEIGHT = 400;

    /**
     * 导出excel通用方法
     *
     * @param table 表格
     */
    static Workbook toExcel(Table table) {
        List<Header> headers = table.getHeaders();
        String tableName = table.getTableName();
        Set<String> conditions = table.getConditions();
        List<Map> baseData = table.getBaseData();
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Drawing<?> patriarch = sheet.createDrawingPatriarch();
        Map<String, CellStyle> cellStyleMap = new HashMap<>();
        CellStyle contentStyle = ExcelStyle.contentStyle(workbook, cellStyleMap);

        // 写表名
        if (StringUtils.isNotBlank(tableName)) {
            Row tableNameRow = sheet.createRow(0);
            tableNameRow.setHeight((short) (ROW_HEIGHT * 2));
            Cell tableNameRowCell = tableNameRow.createCell(0);
            tableNameRowCell.setCellStyle(ExcelStyle.tableNameStyle(workbook, cellStyleMap));
            tableNameRowCell.setCellValue(tableName);
        }

        writeCondition(conditions, sheet); // 写条件

        if (headers == null || headers.size() <= 0) {
            // 没有表头，不再往下执行了
            return workbook;
        }
        writeHeader(headers, sheet, ExcelStyle.headerStyle(workbook, cellStyleMap)); // 写表头

        if (baseData == null || baseData.size() <= 0) {
            // 没有数据，不再往下执行了
            return workbook;
        }
        // 写内容
        int lastRowNum = sheet.getLastRowNum();
        List<String> allKey = getAllKey(headers); // 获取所有取值用的键
        Map<Integer, Integer> maxWidthMap = new HashMap<>(); // 列宽度最大值map，键为列数，值为最大宽度
        for (Map data : baseData) {
            Row row = sheet.createRow(++lastRowNum);
            row.setHeight(ROW_HEIGHT);
            for (int j = 0; j < allKey.size(); j++) {
                String key = allKey.get(j);
                Cell cell = row.createCell(j);
                cell.setCellStyle(contentStyle);
                Object value = data.get(key);
                String valueStr = processValue(value);
                cell.setCellValue(valueStr);

                setCellWidth(maxWidthMap, j, valueStr);
            }
        }

        // 写合计
        Map<String, Object> totalData = table.getTotalData();
        if (totalData != null && totalData.size() > 0) {
            if (!isAllBlank(totalData)) {// 如果合计里全部都是空则不写合计行
                lastRowNum = sheet.getLastRowNum();
                Row row = sheet.createRow(++lastRowNum);
                row.setHeight((short) (ROW_HEIGHT + 50));
                for (int j = 0; j < allKey.size(); j++) {
                    String key = allKey.get(j);
                    Cell cell = row.createCell(j);
                    Object value = table.getTotalData().get(key);
                    String valueStr = processValue(value);
                    if (StringUtils.isBlank(valueStr) && j == 0) {
                        valueStr = "合计";
                    }
                    cell.setCellStyle(ExcelStyle.totalStyle(workbook, cellStyleMap));
                    if (NumberUtils.isNumber(valueStr)) {
                        cell.setCellValue(Double.parseDouble(valueStr));
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellStyle(ExcelStyle.totalNumberStyle(workbook, cellStyleMap));
                    } else {
                        cell.setCellValue(valueStr);
                    }

                    setCellWidth(maxWidthMap, j, valueStr);
                }
            }
        }

        // 表头的宽度也加入宽度自适应的逻辑里
        List<String> allHeaderTitle = getAllHeaderTitle(headers);
        for (int i = 0; i < allHeaderTitle.size(); i++) {
            String title = allHeaderTitle.get(i);
            setCellWidth(maxWidthMap, i, title);
        }

        // 列宽自适应
        for (int i = 0; i < allKey.size(); i++) {
            sheet.setColumnWidth(i, maxWidthMap.get(i));
        }

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, allHeaderTitle.size() - 1)); // 合并单元格
        int titleRowNum;
        if (conditions != null) {
            titleRowNum = 1 + conditions.size() + HeaderUtils.getMaxDeep(headers); // 表头行数
        } else {
            titleRowNum = 1 + HeaderUtils.getMaxDeep(headers); // 表头行数
        }
        sheet.createFreezePane(0, titleRowNum, 0, titleRowNum); // 冻结表头

        return workbook;
    }

    /**
     * 设置单元格宽
     */
    private static void setCellWidth(Map<Integer, Integer> maxWidthMap, int j, String cellValue) {
        int length = cellValue.getBytes().length * 256 + 200;
        // 这里把宽度最大限制到15000
        if (length > 15000) {
            length = 15000;
        }
        maxWidthMap.put(j, Math.max(length, maxWidthMap.get(j) == null ? 0 : maxWidthMap.get(j)));
    }

    /**
     * map里的value是不是全都是空字符串
     *
     * @param totalData map
     * @return true 表示是全部为空字符串，false表示不是
     */
    private static boolean isAllBlank(Map<String, Object> totalData) {
        boolean isAllBlank = true;
        for (Map.Entry<String, Object> entry : totalData.entrySet()) {
            Object value = entry.getValue();
            if (value != null && StringUtils.isNotBlank(processValue(value))) {
                isAllBlank = false;
            }
        }
        return isAllBlank;
    }

    /**
     * 统一处理值
     */
    private static String processValue(Object value) {
        String valueStr = value == null ? "" : value.toString();
        // 去掉小数点后面多余的0
        valueStr = NumberUtils.trimZero(valueStr);
        return valueStr;
    }

    private static List<String> getAllHeaderTitle(List<Header> headers) {
        List<String> allTitle = new ArrayList<>();
        for (Header header : headers) {
            List<Header> child = header.getChild();
            if (child == null || child.size() <= 0) {
                allTitle.add(header.getTitle());
            } else {
                allTitle.addAll(getAllHeaderTitle(child));
            }
        }
        return allTitle;
    }

    private static List<String> getAllKey(List<Header> headers) {
        List<String> allKey = new ArrayList<>();
        for (Header header : headers) {
            List<Header> child = header.getChild();
            if (child == null || child.size() <= 0) {
                allKey.add(header.getKey());
            } else {
                allKey.addAll(getAllKey(child));
            }
        }
        return allKey;
    }

    /**
     * 写条件
     */
    private static void writeCondition(Set<String> conditions, Sheet sheet) {
        if (conditions == null || conditions.size() <= 0) {
            return;
        }
        int lastRowNum = sheet.getLastRowNum();
        for (String condition : conditions) {
            Row row = sheet.createRow(++lastRowNum);
            row.setHeight(ROW_HEIGHT);
            Cell cell = row.createCell(0);
            cell.setCellValue(condition);
        }
    }

    /**
     * 初始化header的父头
     */
    private static void initParentHeader(Header header) {
        List<Header> child = header.getChild();

        for (Header children : child) {
            children.setParent(header);
            List<Header> childrenChild = children.getChild();
            if (childrenChild.size() > 0) {
                initParentHeader(children);
            }
        }
    }

    /**
     * 初始化header的父头
     */
    private static void initParentHeader(List<Header> headers) {
        for (Header header : headers) {
            initParentHeader(header);
        }
    }

    /**
     * 写表头
     */
    private static void writeHeader(List<Header> headers, Sheet sheet, CellStyle cellStyle) {
        initParentHeader(headers); // 保证每个叶子都能找到父节点

        int lastRowNum = sheet.getLastRowNum();
        int startRowNum = lastRowNum + 1;
        List<List<String>> propertyDes = new ArrayList<>();
        getPropertyDes(headers, propertyDes);
        int maxSize = 0;
        for (List<String> tierList : propertyDes) {
            if (tierList.size() > maxSize) {
                maxSize = tierList.size();
            }
        }

        for (List<String> tierList : propertyDes) {
            int size = tierList.size();
            if (size < maxSize) {
                for (int i = 0; i < maxSize - size; i++) {
                    tierList.add("");
                }
            }
        }

        int mergerNum = 0; //合并数

        //给单元格设置值
        for (int i = 0; i < propertyDes.size(); i++) {
            Row row = sheet.createRow(i + startRowNum);
            row.setHeight(ROW_HEIGHT);
            for (int j = 0; j < propertyDes.get(i).size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(propertyDes.get(i).get(j));
            }
        }
        Map<Integer, List<Integer>> map = new HashMap<>();   // 合并行时要跳过的行列
        //合并行
        for (int i = 0; i < propertyDes.get(propertyDes.size() - 1).size(); i++) {
            if ("".equals(propertyDes.get(propertyDes.size() - 1).get(i))) {
                for (int j = propertyDes.size() - 2; j >= 0; j--) {
                    if (!"".equals(propertyDes.get(j).get(i))) {
                        sheet.addMergedRegion(new CellRangeAddress(j + startRowNum, propertyDes.size() - 1 + startRowNum, i, i)); // 合并单元格
                        break;
                    } else {
                        if (map.containsKey(j + startRowNum)) {
                            List<Integer> list = map.get(j + startRowNum);
                            list.add(i);
                            map.put(j + startRowNum, list);
                        } else {
                            List<Integer> list = new ArrayList<Integer>();
                            list.add(i);
                            map.put(j + startRowNum, list);
                        }
                    }
                }
            }
        }
        //合并列
        for (int i = 0; i < propertyDes.size() - 1; i++) {
            for (int j = 0; j < propertyDes.get(i).size(); j++) {
                List<Integer> list = map.get(i + startRowNum);
                if (list == null || !list.contains(j)) {
                    if ("".equals(propertyDes.get(i).get(j))) {
                        mergerNum++;
                        if (mergerNum != 0 && j == (propertyDes.get(i).size() - 1)) {
                            sheet.addMergedRegion(new CellRangeAddress(i + startRowNum, i + startRowNum, j - mergerNum, j)); // 合并单元格
                            mergerNum = 0;
                        }
                    } else {
                        if (mergerNum != 0) {
                            sheet.addMergedRegion(new CellRangeAddress(i + startRowNum, i + startRowNum, j - mergerNum - 1, j - 1)); // 合并单元格
                            mergerNum = 0;
                        }
                    }
                }
            }
        }
    }

    /**
     * 把表头列表转换成以下格式
     * String[][] propertyDes = {
     * {"1","","2","3","","",""},
     * {"1","1","","3","","3",""},
     * {"","","","3","3","3","3"}
     * };
     *
     * @param headers     表头列表
     * @param propertyDes 结果
     */
    private static void getPropertyDes(List<Header> headers, List<List<String>> propertyDes) {
        getPropertyDes1(headers, headers, propertyDes, 0, 0);
    }

    /**
     * 把表头列表转换成以下格式
     * String[][] propertyDes = {
     * {"1","","2","3","","",""},
     * {"1","1","","3","","3",""},
     * {"","","","3","3","3","3"}
     * };
     *
     * @param headers     表头列表
     * @param propertyDes 结果
     * @param columnStart 开始列，从0开始
     * @param tier        层级，从0开始
     */
    private static void getPropertyDes1(List<Header> allHeaders, List<Header> headers, List<List<String>> propertyDes, int columnStart, int tier) {
        for (Header header : headers) {
            List<String> tierList;
            if (tier >= propertyDes.size()) {
                tierList = new ArrayList<>();
                propertyDes.add(tierList);
            } else {
                tierList = propertyDes.get(tier);
            }

            // 在加之前，检查要不要补空字符串
            int preNodeQuantity = 0; // 本节点之前的节点数量
            // 1. 看自己这颗树同一层级前面有多少节点
            Header rootHeader = getRootHeader(header);
            List<Header> HeadersByTier = listHeaderByTier(rootHeader, tier);
            for (Header HeaderByTier : HeadersByTier) {
                // 遍历这些节点，计算本节点之前有几个节点
                if (HeaderByTier == header) {
                    break;
                }
                preNodeQuantity += getMaxChildQuantity(HeaderByTier);
            }

            // 2. 看自己这颗树前面的树的叶子节点数量总和
            for (Header allRootHeader : allHeaders) {
                if (rootHeader == allRootHeader) {
                    break;
                }
                preNodeQuantity += getMaxChildQuantity(allRootHeader);
            }

            if (tierList.size() < preNodeQuantity) {
                // 如果当前层级的数量小于本节点之前的节点数量，用空字符串补
                int size = tierList.size();
                for (int i = 0; i < preNodeQuantity - size; i++) {
                    tierList.add("");
                }
            }

            tierList.add(header.getTitle());
            columnStart++;

            List<Header> child = header.getChild();
            if (child != null && child.size() > 0) {
                getPropertyDes1(allHeaders, child, propertyDes, columnStart, tier + 1);
            }
        }
    }

    /**
     * 获取树某一层的所有节点
     *
     * @param rootHeader 根节点
     * @param targetTier 层级
     * @return 节点列表
     */
    private static List<Header> listHeaderByTier(Header rootHeader, int targetTier) {
        List<Header> Headers = new ArrayList<>();
        if (targetTier <= 0) {
            Headers.add(rootHeader);
            return Headers;
        }

        List<Header> child = rootHeader.getChild();
        if (child.size() <= 0) {
            // 只有一层，目标层级却大于一层，直接返回一个空列表
            return Headers;
        }

        int currentTier = 0; // 当前层级
        listHeaderByTier(child, Headers, targetTier, currentTier + 1);
        return Headers;
    }

    private static void listHeaderByTier(List<Header> child, List<Header> Headers, int targetTier, int currentTier) {
        if (targetTier == currentTier) {
            // 找到目标层级，直接结束递归
            Headers.addAll(child);
            return;
        }

        for (Header Header : child) {
            List<Header> nextTierChild = Header.getChild();
            if (targetTier > currentTier && nextTierChild.size() <= 0) {
                // 如果目标层级比当前层级大，但这个节点已经没有子节点了，也算是同一层的节点
                Headers.add(Header);
                continue;
            }
            listHeaderByTier(nextTierChild, Headers, targetTier, currentTier + 1);
        }
    }

    /**
     * 获取树的根
     *
     * @param Header 树里的某个节点
     * @return 返回最顶级的根节点
     */
    private static Header getRootHeader(Header Header) {
        Header parent = Header.getParent();
        if (parent == null) {
            return Header;
        } else {
            return getRootHeader(parent);
        }
    }

    private static int getMaxChildQuantity(Header header) {
        List<Header> child = header.getChild();
        if (child.size() <= 0) {
            return 1;
        }
        int quantity = 0;
        for (Header Header : child) {
            if (Header.getChild().size() <= 0) {
                ++quantity;
            } else {
                quantity += getMaxChildQuantity(Header);
            }
        }
        return quantity;
    }

    private static int getMaxColumn(List<Header> headers) {
        int maxColumn = 0;
        for (Header header : headers) {
            List<Header> child = header.getChild();
            if (child.size() <= 0) {
                maxColumn += 1;
            } else {
                maxColumn += getMaxChildQuantity(header);
            }
        }
        return maxColumn;
    }

    private static String getCellValue(Cell cell, Workbook wb) {
        String value = "";
        if (cell != null) {

            switch (cell.getCellTypeEnum()) {
                case STRING:
                    value = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        if (date != null) {
                            value = new SimpleDateFormat("yyyy-MM-dd").format(date);
                        } else {
                            value = "";
                        }
                    } else {
                        cell.setCellType(CellType.STRING);
                        value = cell.getStringCellValue() + "";
                    }
                    break;
                case FORMULA:
                    FormulaEvaluator formulaEvaluator;
                    if (wb instanceof HSSFWorkbook) {
                        formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) wb);
                    } else {
                        formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
                    }
                    value = String.valueOf(formulaEvaluator.evaluate(cell).getNumberValue());
                    break;
                default:
                    value = "";
            }
        }
        return value;
    }

    private static void drawPictureInfoExcel(Workbook wb, Drawing<?> patriarch, int rowIndex, int cellIndex, String pictureUrl) {
        //rowIndex代表当前行
        try {
            if (StringUtils.isNotBlank(pictureUrl)) {
                URL url = new URL(pictureUrl);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");
                conn.setConnectTimeout(5 * 1000);
                InputStream inStream = conn.getInputStream();
                byte[] data = readInputStream(inStream);
                ClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 250, (short) cellIndex, rowIndex, (short) cellIndex, rowIndex);
                patriarch.createPicture(anchor, wb.addPicture(data, XSSFWorkbook.PICTURE_TYPE_JPEG));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static byte[] readInputStream(InputStream inStream) throws Exception {
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len = 0;
        while ((len = inStream.read(buffer)) != -1) {
            outStream.write(buffer, 0, len);
        }
        byte[] data = outStream.toByteArray();
        outStream.close();
        inStream.close();
        return data;
    }

}
