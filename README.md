# TreeExcel [![](https://jitpack.io/v/shuangmulin/treeexcel.svg)](https://jitpack.io/#shuangmulin/treeexcel)
##### 基于poi实现Excel多级表头导出方案

## 使用前说明
* poi依赖3.17版本

## 环境依赖
JDK 1.8 +

## 使用步骤
### Maven
1. 引入依赖
```xml
<project>
    <!-- 设置 jitpack.io 仓库 -->
    <repositories>
        <repository>
            <id>jitpack.io</id>
            <url>https://jitpack.io</url>
        </repository>
    </repositories>

    <dependencies>
        <!-- treeexcel -->
        <dependency>
            <groupId>com.github.shuangmulin</groupId>
            <artifactId>treeexcel</artifactId>
            <version>1.0.0</version>
        </dependency>
        
        <!-- poi -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.17</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.17</version>
        </dependency>
    </dependencies>
</project>
```

2. 使用代码demo
```java
public class Test{

    public static void main(String[] args) throws IOException {
        // 注意这里的`#`这个符号，是用来分割表头层级的，只需要按这个格式写，执行后相同的父级会自动合并（看下面的效果图）
        Table table = Table.builder()
                .tableName("销售汇总报表")
                .addHeader("styleNo", "商品数据#商品档案")
                .addHeader("styleName", "商品数据#商品名称")
                .addHeader("color", "颜色")
                .addHeader("S", "统计数据#尺码明细#S")
                .addHeader("M", "统计数据#尺码明细#M")
                .addHeader("L", "统计数据#尺码明细#L")
                .addHeader("saleQuantity", "统计数据#销售总数")
                .addHeader("saleAmount", "统计数据#销售金额")
                .baseData(getData())
                .build();
        Workbook workbook = TreeExcel.getExcel(table);
        workbook.write(new FileOutputStream("C:\\销售汇总报表.xlsx")); // 自行更换导出路径
    }

    public static List<Map> getData() {
        List<Map> baseData = new ArrayList<>();
        Map<String,Object> map= new HashMap<>();
        map.put("styleNo", "SP001");
        map.put("styleName", "毛领羽绒");
        map.put("color", "红色");
        map.put("S", 10);
        map.put("M", 22);
        map.put("L", 998);
        map.put("saleQuantity", 998);
        map.put("saleAmount", 10000);
        baseData.add(map);

        map= new HashMap<>();
        map.put("styleNo", "SP001");
        map.put("styleName", "毛领羽绒");
        map.put("color", "绿色");
        map.put("S", 11);
        map.put("M", 33);
        map.put("L", 44);
        map.put("saleQuantity", 12);
        map.put("saleAmount", 323);
        baseData.add(map);

        map= new HashMap<>();
        map.put("styleNo", "SP001");
        map.put("styleName", "毛领羽绒");
        map.put("color", "黑色");
        map.put("S", 12);
        map.put("M", 43);
        map.put("L", 11);
        map.put("saleQuantity", 3123);
        map.put("saleAmount", 123123);
        baseData.add(map);

        map= new HashMap<>();
        map.put("styleNo", "SP002");
        map.put("styleName", "真丝上衣");
        map.put("color", "红色");
        map.put("S", 123);
        map.put("M", 34);
        map.put("L", 53);
        map.put("saleQuantity", 424);
        map.put("saleAmount", 756);
        baseData.add(map);

        map= new HashMap<>();
        map.put("styleNo", "SP002");
        map.put("styleName", "真丝上衣");
        map.put("color", "红色");
        map.put("S", 4);
        map.put("M", 23);
        map.put("L", 122);
        map.put("saleQuantity", 342);
        map.put("saleAmount", 5345);
        baseData.add(map);
        return baseData;
    }
    
}
```
3. 效果
以上代码执行，就会得到一张这样的Excel

<img alt="treeexcel" src="https://raw.githubusercontent.com/shuangmulin/treeexcel/master/img/img.png">
