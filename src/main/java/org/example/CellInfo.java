package org.example;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @Description: TODO
 * @author: scott
 * @date: 2020年10月05日 20:26
 */

@Data
public class CellInfo {
    /** 样式 */
    CellStyle cellStyle;
    /** 合并单元格 */
    MergedResult mergedResult;
    /** 行高 */
    short rowHeigth;
    /** 列宽 */
    int columnWidth;
    /** 数据之前显示的内容 */
    private String initially;
    /** 数据的名字 */
    private String name;
    /** 数据之后显示的内容 */
    private String ending;
    /** 数据所在行 */
    private Integer rowIndex;
    /** 数据所在裂 */
    private Integer cellIndex;

    public CellInfo(){}

    public CellInfo(String initially, String name, String ending, Integer rowIndex, Integer cellIndex) {
        this.initially = initially;
        this.name = name;
        this.ending = ending;
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
    }

    public CellInfo(CellStyle cellStyle, short rowHeigth, int columnWidth, String initially,
                    String name, String ending, Integer rowIndex, Integer cellIndex, MergedResult mergedResult) {
        this.cellStyle = cellStyle;
        this.rowHeigth = rowHeigth;
        this.columnWidth = columnWidth;
        this.initially = initially;
        this.name = name;
        this.ending = ending;
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
        this.mergedResult = mergedResult;
    }
}
