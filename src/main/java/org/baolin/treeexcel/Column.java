package org.baolin.treeexcel;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * description: 列信息
 *
 * @author baolin
 **/
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Column {

    /**
     * key
     */
    private String key;
    /**
     * 名称
     */
    private String title;

}
