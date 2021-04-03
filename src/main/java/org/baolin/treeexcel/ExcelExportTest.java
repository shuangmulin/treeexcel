package org.baolin.treeexcel;

import org.baolin.treeexcel.utils.HeaderUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author 钟宝林
 **/
class ExcelExportTest {

    public static void main(String[] args) throws IOException {
//        List<Map> baseDatas = new ArrayList<>();
//
//        for (int i = 0; i < 200; i++) {
//            baseDatas.add(getData(i));
//        }
//
//        Table table = new Table();
//        table.setTableName("测试报表");
//        table.setConditions(Collections.singleton("[商品款号：fsdf、法师答复、fasdf]"));
//        table.setBaseDatas(baseDatas);
//        table.setHeaders(getHeaders());
//        Workbook workbook = ExcelUtils.toExcel(table);
//        workbook.write(new FileOutputStream(new File("/Users/regent/Desktop/file.xlsx")));
        System.out.println(HeaderUtils.getMaxDeep(getHeaders()));
    }

    private static Map<String, Object> getData(int i) {
        Map<String, Object> row1 = new HashMap<>();
        row1.put("unitName", "0");
        row1.put("materialDetail", "a222222222" + i);
        row1.put("materialNo", "33333s333333" + i);
        row1.put("materialKindName", "4444d44444444" + i);
        row1.put("picture", "5555555555f555" + i);
        row1.put("materialName", "666g66666666" + i);
        row1.put("a", "77777777g7777" + i);
        row1.put("b", "8888888d88" + i);
        row1.put("c", "9999999s9999" + i);
        row1.put("d", "0000000s000000" + i);
        return row1;
    }

    private static List<Header> getHeaders() {
        List<Header> headers = new ArrayList<>();
        Header header = new Header("组织工厂", "unitName");
        Header header4 = new Header("fasdfdsaf", "unitName");
        headers.add(header4);
        Header header2 = new Header("物料明细物料明细物料明细", "materialDetail");
        header2.addChild(new Header("物料款号", "materialNo")

                .addChild(new Header("图片", "picture"))
                .addChild(new Header("物料", "materialName")
                        .addChild(new Header("最里面的", "a")
                                .addChild(new Header("最里面的1", "b"))
                                .addChild(new Header("最里面的2最里面的2最里面的2", "c"))
                        )
                ).addChild(new Header("分类", "materialKindName")));
        headers.add(header);
        headers.add(header2);
        Header header3 = new Header("组织工厂33", "unitName");
        headers.add(header3);
        return headers;
    }
}
