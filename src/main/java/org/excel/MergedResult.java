package org.excel;

import lombok.Data;

/**
 * @author Administrator
 */
@Data
public class MergedResult {
    /** 是否合并单元格 */
    private boolean isMerged;
    /** 行下标 */
    private int rowIndex;
    /** 列下标 */
    private int columnIndex;
    /** 合并的行 开始下标 */
    private int firstRow;
    /** 合并的行 结束下标 */
    private int lastRow;
    /** 合并的列 开始下标 */
    private int firstColumn;
    /** 合并的列 结束下标 */
    private int lastColumn;
    /** 合并的行数 */
    private int rowMergeNum;
    /** 合并的列数 */
    private int columnMergeNum;

}
