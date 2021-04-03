package org.baolin.treeexcel;

import org.baolin.treeexcel.utils.HeaderUtils;
import org.baolin.treeexcel.utils.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 一个Excel表格
 *
 * @author 钟宝林
 **/

public class Table {

    /**
     * 表格名称
     */
    private String tableName;
    /**
     * 条件文案
     */
    private Set<String> conditions;
    /**
     * 表头
     */
    private List<Header> headers;
    /**
     * 合计数据
     */
    private Map<String, Object> totalData;
    /**
     * 表数据
     */
    private List<Map> baseData;

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public Set<String> getConditions() {
        return conditions;
    }

    public void setConditions(Set<String> conditions) {
        this.conditions = conditions;
    }

    public List<Header> getHeaders() {
        return headers;
    }

    public void setHeaders(List<Header> headers) {
        this.headers = headers;
    }

    public Map<String, Object> getTotalData() {
        return totalData;
    }

    public void setTotalData(Map<String, Object> totalData) {
        this.totalData = totalData;
    }

    public List<Map> getBaseData() {
        return baseData;
    }

    public void setBaseData(List<Map> baseData) {
        this.baseData = baseData;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private String tableName;
        private Set<String> conditions;
        private List<Header> headers;
        private Map<String, Object> totalData;
        private List<Map> baseData;
        private List<Column> tempColumns;

        private Builder() {
        }

        public Builder tableName(String tableName) {
            this.tableName = tableName;
            return this;
        }

        public Builder conditions(Set<String> conditions) {
            this.conditions = conditions;
            return this;
        }

        public Builder headers(List<Header> headers) {
            this.headers = headers;
            return this;
        }

        public Builder headersWithCol(List<Column> columns) {
            this.headers = columns == null || columns.size() <= 0 ? null : HeaderUtils.getHeaders(columns);
            return this;
        }

        public Builder addHeader(String key, String title) {
            if (StringUtils.isBlank(key) || StringUtils.isBlank(title)) {
                return this;
            }
            this.tempColumns = this.tempColumns == null ? new ArrayList<>() : this.tempColumns;
            this.tempColumns.add(new Column(key, title));
            return this;
        }

        public Builder headersWithMap(Map<String, String> headers) {
            if (headers != null && headers.size() > 0) {
                List<Column> columns = new ArrayList<>();
                for (Map.Entry<String, String> entry : headers.entrySet()) {
                    String key = entry.getKey();
                    String value = entry.getValue();
                    columns.add(new Column(key, value));
                }
                this.headers = HeaderUtils.getHeaders(columns);
            }
            return this;
        }

        public Builder totalData(Map<String, Object> totalData) {
            this.totalData = totalData;
            return this;
        }

        public Builder baseData(List<Map> baseData) {
            this.baseData = baseData;
            return this;
        }

        public Table build() {
            if (this.tempColumns != null && this.tempColumns.size() > 0) {
                this.headers = HeaderUtils.getHeaders(this.tempColumns);
            }

            Table table = new Table();
            table.setTableName(tableName);
            table.setConditions(conditions);
            table.setHeaders(headers);
            table.setTotalData(totalData);
            table.setBaseData(baseData);
            return table;
        }
    }
}
