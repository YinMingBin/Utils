package org.example;

import lombok.Data;

/**
 * @Description: TODO
 * @author: scott
 * @date: 2020年09月27日 21:08
 */

@Data
public class CellContent {
    private String initially;
    private String name;
    private String ending;
    private Integer rowSite;
    private Integer cellSite;

    public CellContent(String initially, String name, String ending, Integer rowSite, Integer cellSite) {
        this.initially = initially;
        this.name = name;
        this.ending = ending;
        this.rowSite = rowSite;
        this.cellSite = cellSite;
    }
}
