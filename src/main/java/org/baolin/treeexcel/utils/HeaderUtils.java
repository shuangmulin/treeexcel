package org.baolin.treeexcel.utils;

import org.baolin.treeexcel.Column;
import org.baolin.treeexcel.Header;

import java.util.*;

/**
 * 头相关的工具
 *
 * @author 钟宝林
 **/
public class HeaderUtils {

    /**
     * 遍历表头（树）
     *
     * @param headers        表头（树）
     * @param headerCallback 回调
     */
    public static void forEachHeader(List<Header> headers, HeaderCallback headerCallback) {
        for (Header header : headers) {
            headerCallback.callback(header);
            List<Header> child = header.getChild();
            if (child != null && child.size() > 0) {
                forEachHeader(child, headerCallback);
            }
        }
    }


    public interface HeaderCallback {
        void callback(Header header);
    }

    /**
     * 获取树的最大深度
     *
     * @param headers 树列表
     * @return 最大深度
     */
    public static int getMaxDeep(List<Header> headers) {
        int max = 0;
        for (Header header : headers) {
            int maxDeep = getMaxDeep(header);
            if (max < maxDeep) {
                max = maxDeep;
            }
        }
        return max;
    }

    /**
     * 获取树的最大深度
     *
     * @param header 树
     * @return 最大深度
     */
    public static int getMaxDeep(Header header) {
        List<Header> child = header.getChild();
        if (child != null && child.size() > 0) {
            int max = 0;
            for (Header Header : child) {
                int maxDeep = getMaxDeep(Header);
                if (max < maxDeep) {
                    max = maxDeep;
                }
            }
            return max + 1;
        } else {
            return 1;
        }
    }

    /**
     * 用数据库列转换成有结构的表头
     * 约定：
     * > {@link Column#getKey()}对应{@link Header#getKey()}
     * > {@link Column#getTitle()}对应{@link Header#getTitle()}，格式为：一级名称#二级名称#三级名称...，会转换成对应结构的Header
     * 可以指定图片类型，格式为：物料图片->picture
     *
     * @param columns 数据库列
     * @return {@link Header} 列表
     */
    public static List<Header> getHeaders(List<Column> columns) {
        return getHeaders(columns, false);
    }

    public static List<Header> getHeaders(List<Column> columns, boolean isMergeChild) {
        List<Header> headers = new ArrayList<>();
        for (Column column : columns) {
            String key = column.getKey();
            String columnTitle = column.getTitle();

            String[] split = columnTitle.split("#");

            Header header = new Header();
            header.setTitle(split[split.length - 1]);
            header.setKey(key);
            String title = header.getTitle();
            String[] strategySplit = title.split("->");
            if (strategySplit.length > 1) {
                String strategy = strategySplit[1];
                if ("picture".equals(strategy)) {
                    header.setType(Header.Type.PICTURE);
                }
            }

            if (split.length > 1) {
                // 非顶级Header
                List<Header> parentChild = headers;
                // 父
                for (int i = 0; i < split.length; i++) {
                    String parentTitle = split[i];
                    Header parentHeader;
                    if (isMergeChild) {
                        parentHeader = getByTitle(parentTitle, parentChild);
                    } else {
                        parentHeader = getParentByTitle(parentTitle, parentChild);
                    }
                    if (parentHeader == null) {
                        if (i == split.length - 1) {
                            // 最后一层
                            parentHeader = header;
                        } else {
                            parentHeader = new Header();
                            parentHeader.setTitle(parentTitle);
                        }
                        parentChild.add(parentHeader);
                    }
                    parentChild = parentHeader.getChild();
                }
            } else {
                headers.add(header);
            }
        }

        return headers;
    }

    private static Header getParentByTitle(String parentTitle, List<Header> parentChild) {
        if (parentChild == null || parentChild.size() <= 0) {
            return null;
        }
        Header header = parentChild.get(parentChild.size() - 1);
        if (Objects.equals(header.getTitle(), parentTitle)) {
            return header;
        }
        return null;
    }

    /**
     * 把数据库动态列转化成树形结构表头
     *
     * @param columns       数据库列，比如：
     *                      出货数#$s1
     *                      出货数#$s2
     * @param dynamicHeader 动态列，结构例子：
     *                      <尺码组ID， <结果集占位符列名，尺码名称>>
     *                      <code>
     *                      {
     *                      "尺码组1":{
     *                      "s1": "S",
     *                      "s2": "M"
     *                      },
     *                      "尺码组2":{
     *                      "s1": "20",
     *                      "s2": "30"
     *                      }
     *                      }
     *                      <code/>
     * @return 处理逻辑为：出货数#$s1 -> 出货数#S#20，出货数#$s2 -> 出货数#M#30
     */
    public static List<Header> getHeaders(List<Column> columns, LinkedHashMap<String, LinkedHashMap<String, String>> dynamicHeader, boolean isMergeChild) {
        if (dynamicHeader == null || dynamicHeader.size() <= 0) {
            return getHeaders(columns, isMergeChild);
        }
        Collection<LinkedHashMap<String, String>> dynamicHeaderList = dynamicHeader.values();
        Map<String, List<String>> headerTier = new HashMap<>();
        List<String> keySet = new ArrayList<>();
        for (Map<String, String> headerMap : dynamicHeaderList) {
            for (Map.Entry<String, String> entry : headerMap.entrySet()) {
                String key = entry.getKey();
                if (keySet.contains(key)) {
                    continue;
                }
                keySet.add(key);
            }
        }

        // 长度不够，用空格占位
        for (String key : keySet) {
            for (Map<String, String> headerMap : dynamicHeaderList) {
                String value = headerMap.get(key);
                value = value == null ? " " : value;
                headerTier.computeIfAbsent(key, k -> new ArrayList<>()).add(value);
            }
        }

        for (Column column : columns) {
            String columnComment = column.getTitle();

            String[] split = columnComment.split("#");

            List<String> columnCommentNew = new ArrayList<>();
            for (String title : split) {
                if (title.startsWith("$")) {
                    // 动态表头
                    String key = title.replaceAll("\\$", "");
                    List<String> headers = headerTier.get(key);
                    StringBuilder titleNew = new StringBuilder();
                    for (String header : headers) {
                        titleNew.append(header).append("#");
                    }
                    titleNew.deleteCharAt(titleNew.length() - 1);
                    title = titleNew.toString();
                }
                columnCommentNew.add(title);
            }
            column.setTitle(StringUtils.join(columnCommentNew, "#"));
        }

        return getHeaders(columns, isMergeChild);
    }

    private static Header getByTitle(String title, List<Header> headers) {
        for (Header header : headers) {
            if (title.equals(header.getTitle())) {
                return header;
            }
        }
        return null;
    }

}
