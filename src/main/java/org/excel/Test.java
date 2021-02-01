package org.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.*;

/**
 * 测试
 * @author Administrator
 */
public class Test {
    public static void main(String[] args) throws IOException {
        ExcelUtil export = new ExcelUtil();
        export.excelAdaptive(new HSSFWorkbook(), new HashMap<>(1));
        Workbook sheets = ExcelUtil.excelAdaptive("", new HashMap<>(1));
        Sheet sheetAt = sheets.getSheetAt(0);
        ExcelUtil.splitMergedRegion(sheetAt, 0, 0, 0, 0);
        ExcelUtil.copySheet(sheetAt, null);
        export.empty();
        System.out.println(sheets);
    }

}
