package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class Test {
    public static void main(String[] args) throws IOException {
        ExcelExport export = new ExcelExport();
        export.excelAdaptive(new HSSFWorkbook(), new HashMap<>(1));
        Workbook sheets = ExcelExport.excelAdaptive("", new HashMap<>(1));
        ExcelExport.splitMergedRegion(null, 0, 0, 0, 0);
        ExcelExport.copySheet(null, null);
        export.empty();
    }

}
