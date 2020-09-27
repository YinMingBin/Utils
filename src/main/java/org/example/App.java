package org.example;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Hello world!
 */
public class App {
    private Map<String, CellContent> all = new HashMap<>();
    private final String start = "${";
    private final String finish = "}";

    public static void main(String[] args) throws IOException {
        System.out.println("Hello World!");
        InputStream is = new FileInputStream("C:/Users/Y2753/Desktop/cTest.xls");
        POIFSFileSystem ps = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(ps);
        Map<String, Object> datas = new HashMap<>();
        datas.put("productNo", "001");
        datas.put("iqcNo", "002");
        SimpleDateFormat sdf = new SimpleDateFormat("YYYY-MM-DD HH:mm:ss");
        datas.put("createTime", sdf.format(new Date()));
        datas.put("inspectorName", "尹明彬");

        App app = new App();
        app.excel(wb, datas);
        FileOutputStream os = new FileOutputStream("D:/A临时/excel/test.xls");
        wb.write(os);
        os.flush();
        os.close();
    }

    public void excel(HSSFWorkbook wb, Map<String, Object> datas) {
        HSSFSheet sheetAt = wb.getSheetAt(0);
        initialize(sheetAt, datas);
        setBasicData(sheetAt, datas);

        HSSFSheet sheet = wb.createSheet();
        CopySheetUtil.copySheets(sheet, sheetAt);


    }

    public void initialize(HSSFSheet sheetAt, Map<String, Object> datas){
        int firstRowNum = sheetAt.getFirstRowNum();
        int lastRowNum = sheetAt.getLastRowNum();
        for(int i = firstRowNum; i < lastRowNum; i++){
            HSSFRow row = sheetAt.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();
            for(int j = firstCellNum; j < lastCellNum; j++){
                HSSFCell cell = row.getCell(j);
                String cellValue = cell.toString();
                int initially = cellValue.indexOf(this.start);
                if(initially >=0){
                    int ending = cellValue.indexOf(this.finish);
                    String substring = cellValue.substring(initially+this.start.length(), ending);
                    String[] split = substring.split("\\.");
                    int splitLen = split.length;
                    String s = split[0];
                    Object o = datas.get(s);
                    if(o != null){
                        if(!(o instanceof List || o instanceof Map)){
                            this.all.put(s, new CellContent(cellValue.substring(0, initially), s, cellValue.substring(ending+this.finish.length()), i, j));
                        }else if(splitLen >= 2) {
                            this.all.put(s, new CellContent(cellValue.substring(0, initially), split[1], cellValue.substring(ending+this.finish.length()), i, j));
                        }else{
                            this.all.put(s, new CellContent(cellValue.substring(0, initially), s, cellValue.substring(ending+this.finish.length()), i, j));
                        }
                    }
                }
            }
        }
    }

    public void setBasicData(HSSFSheet sheetAt, Map<String, Object> datas){
        for(Map.Entry<String, Object> data : datas.entrySet()){
            Object value = data.getValue();
            if(!(value instanceof List || value instanceof Map)){
                String key = data.getKey();
                CellContent cellContent = all.get(key);
                if(cellContent != null){
                    HSSFRow row = sheetAt.getRow(cellContent.getRowSite());
                    HSSFCell cell = row.getCell(cellContent.getCellSite());
                    String initially = cellContent.getInitially();
                    String ending = cellContent.getEnding();
                    cell.setCellValue(initially+value+ending);
                }
            }
        }
    }

}
