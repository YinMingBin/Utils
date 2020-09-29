package org.example;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Hello world!
 */
public class App {
    private Map<String, CellContent> all = new HashMap<>();
    private Map<String, List<CellContent>> allArray = new HashMap<>();
    private final String start = "${";
    private final String finish = "}";

    public static void main(String[] args) throws IOException {
//        ClassPathResource resource = new ClassPathResource("test.xls");
        InputStream is = new FileInputStream("target/classes/test.xls");
        POIFSFileSystem ps = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(ps);
        Map<String, Object> datas = new HashMap<>();
        datas.put("productNo", "001");
        datas.put("iqcNo", "002");
        SimpleDateFormat sdf = new SimpleDateFormat("YYYY-MM-DD HH:mm:ss");
        datas.put("createTime", sdf.format(new Date()));
        datas.put("inspectorName", "尹明彬");
        List<Map<String, Object>> paramRecords = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            Map<String, Object> paramRecordMap = new HashMap();
            paramRecordMap.put("inspectionTypeName", "iTN"+i);
            paramRecordMap.put("inspectionName", "iN"+i);
            paramRecordMap.put("chkDevName", "cDN"+i);
            paramRecordMap.put("prodUnit", "pU"+i);
            paramRecordMap.put("standard", "standard"+i);
            paramRecordMap.put("sl", "sl"+i);
            paramRecordMap.put("usl", "usl"+i);
            paramRecordMap.put("lsl", "lsl"+i);
            paramRecordMap.put("result", "result"+i);
            paramRecords.add(paramRecordMap);
        }
        datas.put("X-paramRecords", paramRecords);

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

        Map<String, Object> xDatas = new HashMap<>();
        datas.forEach((key, value) -> {
            if(key.startsWith("X-")){
                xDatas.put(key, value);
            }
        });
        setXAllArray(sheet, xDatas);

    }

    public void initialize(HSSFSheet sheet, Map<String, Object> datas){
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for(int i = firstRowNum; i < lastRowNum; i++){
            HSSFRow row = sheet.getRow(i);
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
                        if(!(o instanceof List || o instanceof Map || o.getClass().isArray())){
                            this.all.put(s, new CellContent(cellValue.substring(0, initially), s, cellValue.substring(ending+this.finish.length()), i, j));
                        }else{
                            List<CellContent> cellContents = this.allArray.get(s);
                            if(CollectionUtils.isEmpty(cellContents)){
                                cellContents = new ArrayList<>();
                            }
                            if(splitLen >= 2) {
                                cellContents.add(new CellContent(cellValue.substring(0, initially), split[1], cellValue.substring(ending + this.finish.length()), i, j));
                            }else{
                                cellContents.add(new CellContent(cellValue.substring(0, initially), s, cellValue.substring(ending+this.finish.length()), i, j));
                            }
                            this.allArray.put(s, cellContents);
                        }
                    }
                }
            }
        }
    }

    public void setBasicData(HSSFSheet sheet, Map<String, Object> datas){
        datas.forEach((key, value) -> {
            if (!(value instanceof List || value instanceof Map)) {
                CellContent cellContent = all.get(key);
                if (cellContent != null) {
                    HSSFRow row = sheet.getRow(cellContent.getRowSite());
                    HSSFCell cell = row.getCell(cellContent.getCellSite());
                    String initially = cellContent.getInitially();
                    String ending = cellContent.getEnding();
                    cell.setCellValue(initially + value + ending);
                }
            }
        });
    }

    public void setXAllArray(HSSFSheet sheet, Map<String, Object> datas){
        datas.forEach((key, value) -> {
            List<CellContent> cellContents = allArray.get(key);
            if(!CollectionUtils.isEmpty(cellContents)){
                if(value instanceof Map){
                    setXMapData((HashMap) value, cellContents, sheet, -10000);
                }else if(value instanceof List){
                    List<Object> list = (List) value;
                    int i = 0;
                    for (Object o : list) {
                        if(o instanceof Map){
                            setXMapData((HashMap) o,cellContents, sheet, i);
                        }
                        i++;
                    }
                }else{
                    Object[] objs = (Object[]) value;
                    int i = 0;
                    for (Object obj : objs) {
                        String s = obj.toString();
                        for (CellContent cellContent : cellContents) {
                            if(cellContent.getName().equals(s)){
                                Integer rowSite = cellContent.getRowSite();
                                Integer cellSite = cellContent.getCellSite();
                                HSSFCell hssfCell = setCellValue(s, sheet, rowSite, cellSite + i);
                                HSSFRow row = sheet.getRow(rowSite);
                                HSSFCell cell = row.getCell(cellSite);
                                copyCellStyle(cell, hssfCell);
                            }
                        }
                        i++;
                    }
                }
            }
        });
    }

    public HSSFCell setCellValue(String data, HSSFSheet sheet, int rowSite, int cellSite){
        HSSFRow row = sheet.getRow(rowSite);
        HSSFCell cell = row.getCell(cellSite);
        if(cell == null){
            cell = row.createCell(cellSite);
        }else if(!StringUtils.isEmpty(cell.toString()) && cell.toString().indexOf("${") < 0 ){
            moveCellX(row, cellSite, 1);
        }
        cell.setCellValue(data);
        return cell;
    }

    public void setXMapData(Map<String, Object> datas, List<CellContent> cellContents, HSSFSheet sheet, int site){
        int i = 0;
        for (Map.Entry<String, Object> data : datas.entrySet()){
            String k = data.getKey();
            String v = data.getValue().toString();
            for (CellContent cellContent : cellContents) {
                if(cellContent.getName().equals(k)){
                    Integer rowSite = cellContent.getRowSite();
                    Integer cellSite = cellContent.getCellSite();
                    HSSFCell hssfCell = setCellValue(v, sheet, rowSite, cellSite + (site > -9999 ? site : i));
                    HSSFRow row = sheet.getRow(rowSite);
                    HSSFCell cell = row.getCell(cellSite);
                    copyCellStyle(cell, hssfCell);
                }
            }
            i++;
        }
    }

    public void moveCellX(HSSFRow row, int target,int time){
        HSSFCell cell = row.getCell(target);
        for (int i = 0; i < time; i++) {
            HSSFCell cell1 = row.getCell(target + i + 1);
            if(cell1 == null){
                cell1 = row.createCell(target + 1);
                CopySheetUtil.copyCell(cell, cell1, new HashMap<>());
                copyCellStyle(cell, cell1);
            }else if(StringUtils.isEmpty(cell1.toString())){
                CopySheetUtil.copyCell(cell, cell1, new HashMap<>());
                copyCellStyle(cell, cell1);
            }else{
                moveCellX(row, target + i +1, time);
                CopySheetUtil.copyCell(cell, cell1, new HashMap<>());
            }
        }
    }

    public void copyCellStyle(HSSFCell cell, HSSFCell newCell){
        HSSFCellStyle cellStyle = cell.getCellStyle();
        if(cellStyle != null){
            newCell.setCellStyle(cellStyle);
        }
    }

}
