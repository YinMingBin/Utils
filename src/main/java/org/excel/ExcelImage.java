package org.excel;

import lombok.Data;

/**
 * excel导出图片
 * @author YinMingBin
 */
@Data
public class ExcelImage {
    private byte[] bytes;
    private int startX;
    private int startY;
    private int endX;
    private int endY;

    public ExcelImage(byte[] bytes) {
        this.bytes = bytes;
        this.startX = 0;
        this.startY = 0;
        this.endX = 255;
        this.endY = 255;
    }

    public ExcelImage(byte[] bytes, int startX, int startY, int endX, int endY) {
        this.bytes = bytes;
        this.startX = startX;
        this.startY = startY;
        this.endX = endX;
        this.endY = endY;
    }
}
