package org.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 单元格信息
 * @author YinMingBin
 */
@Data
public class CellInfo {
    /** 样式 */
    CellStyle cellStyle;
    /** 合并单元格 */
    MergedResult mergedResult;
    /** 行高 */
    short rowHeight;
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

    public CellInfo(String initially, String name, String ending, Integer rowIndex, Integer cellIndex) {
        this.initially = initially;
        this.name = name;
        this.ending = ending;
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
    }
}
