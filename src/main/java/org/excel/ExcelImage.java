package org.excel;

import lombok.Data;

/**
 * excel导出图片
 * @author YinMingBin
 */
@Data
public class ExcelImage {
    byte[] bytes;

    public ExcelImage(byte[] bytes) {
        this.bytes = bytes;
    }
}
