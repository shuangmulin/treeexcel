package org.baolin.treeexcel;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * 表头DTO
 *
 * @author 钟宝林
 **/
public class Header implements Serializable {
    private static final long serialVersionUID = -8643938683283748579L;

    /**
     * 文本
     */
    private String title;
    /**
     * 取数据的键
     */
    private String key;
    /**
     * 子头
     */
    private List<Header> child = new ArrayList<>();
    /**
     * 父头
     */
    private Header parent;
    /**
     * 类型
     */
    private Type type = Type.NORMAL;
    /**
     * 排序
     */
    private int sort = 1000;

    public Header(String title, String key) {
        this.title = title;
        this.key = key;
    }

    public Header(String title) {
        this.title = title;
    }

    public Header() {
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public List<Header> getChild() {
        return child;
    }

    public void setChild(List<Header> child) {
        this.child = child;
    }

    public Header addChild(Header header) {
        this.child.add(header);
        return this;
    }

    public Type getType() {
        return type;
    }

    public void setType(Type type) {
        this.type = type;
    }

    public Header getParent() {
        return parent;
    }

    public void setParent(Header parent) {
        this.parent = parent;
    }

    public int getSort() {
        return sort;
    }

    public void setSort(int sort) {
        this.sort = sort;
    }

    /**
     * 类型
     */
    public enum Type {
        NORMAL, PICTURE
    }
}
