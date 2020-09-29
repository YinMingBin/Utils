package org.example;

import lombok.Data;

/**
 * @Description: TODO
 * @author: scott
 * @date: 2020年09月27日 21:08
 */

@Data
public class CellContent {
    /** 数据之前显示的内容 */
    private String initially;
    /** 数据的名字 */
    private String name;
    /** 数据之后显示的内容 */
    private String ending;
    /** 数据所在行 */
    private Integer rowSite;
    /** 数据所在裂 */
    private Integer cellSite;

    public CellContent(String initially, String name, String ending, Integer rowSite, Integer cellSite) {
        this.initially = initially;
        this.name = name;
        this.ending = ending;
        this.rowSite = rowSite;
        this.cellSite = cellSite;
    }
}
