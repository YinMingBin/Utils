package org.example;

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
        ExcelExport export = new ExcelExport();
        export.excelAdaptive(new HSSFWorkbook(), new HashMap<>(1));
        Workbook sheets = ExcelExport.excelAdaptive("", new HashMap<>(1));
        Sheet sheetAt = sheets.getSheetAt(0);
        ExcelExport.splitMergedRegion(sheetAt, 0, 0, 0, 0);
        ExcelExport.copySheet(sheetAt, null);
        export.empty();
        System.out.println(sheets);
    }

}
