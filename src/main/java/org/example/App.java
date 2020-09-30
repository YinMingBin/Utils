package org.example;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
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
    private Map<String, HSSFCellStyle> allCellStyle = new HashMap<>();
    private final String start = "${";
    private final String finish = "}";
    private final String x = "X-";
    private final String y = "Y-";

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

            List<Integer> no = new ArrayList<>();
            for(int j = i/2; j < 10; j++){
                no.add(j);
            }

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
            paramRecordMap.put("Y-no", no);
            paramRecords.add(paramRecordMap);
        }
        datas.put("X-paramRecords", paramRecords);

        App app = new App();
        app.excel(wb, datas);
        FileOutputStream os = new FileOutputStream("C:/Users/Administrator/Desktop/test.xls");
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

        setAllArray(sheet, datas);

    }

    public void initialize(HSSFSheet sheet, Map<String, Object> datas){
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for(int i = firstRowNum; i <= lastRowNum; i++){
            HSSFRow row = sheet.getRow(i);
            if(row != null) {
                short firstCellNum = row.getFirstCellNum();
                short lastCellNum = row.getLastCellNum();
                for (int j = firstCellNum; j < lastCellNum; j++) {
                    HSSFCell cell = row.getCell(j);
                    if(cell != null) {
                        String cellValue = cell.toString();
                        int initially = cellValue.indexOf(this.start);
                        if (initially >= 0) {
                            int ending = cellValue.indexOf(this.finish);
                            String substring = cellValue.substring(initially + this.start.length(), ending);
                            String[] split = substring.split("\\.");
                            int splitLen = split.length;
                            String s = split[0];
                            Object o = datas.get(s);
                            if (o != null) {
                                String subEnding = cellValue.substring(ending + this.finish.length());
                                String str = split[splitLen - 1];
                                if (o instanceof Map || o instanceof List || o.getClass().isArray()) {
                                    List<CellContent> cellContents = this.allArray.get(s);
                                    if (CollectionUtils.isEmpty(cellContents)) {
                                        cellContents = new ArrayList<>();
                                    }
                                    if (splitLen >= 2) {
                                        cellContents.add(new CellContent(cellValue.substring(0, initially), str, subEnding, i, j));
                                    } else {
                                        cellContents.add(new CellContent(cellValue.substring(0, initially), str, subEnding, i, j));
                                    }
                                    this.allArray.put(s, cellContents);
                                    this.allCellStyle.put(str, cell.getCellStyle());
                                } else {
                                    this.all.put(s, new CellContent(cellValue.substring(0, initially), str, subEnding, i, j));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public void setBasicData(HSSFSheet sheet, Map<String, Object> datas){
        all.forEach((key, value) -> {
            Object o = datas.get(key);
            HSSFRow row = sheet.getRow(value.getRowSite());
            HSSFCell cell = row.getCell(value.getCellSite());
            String initially = value.getInitially();
            String ending = value.getEnding();
            cell.setCellValue(initially + o.toString() + ending);
        });
    }

    public void setAllArray(HSSFSheet sheet, Map<String, Object> datas){
        allArray.forEach((key, value) -> {
            Object o = datas.get(key);
            CellContent cellContent = value.get(0);
            eachTransferStop(key, o, value, sheet, cellContent.getRowSite(), cellContent.getCellSite(), -10000);
        });
    }

    public HSSFCell setCellValue(Object data, HSSFSheet sheet, int rowSite, int cellSite){
        return setCellValue(data, sheet, rowSite, cellSite, false);
    }

    public HSSFCell setCellValue(Object data, HSSFSheet sheet, int rowSite, int cellSite, boolean isY){
        HSSFRow row = sheet.getRow(rowSite);
        if(row == null){
            row = sheet.createRow(rowSite);
        }
        HSSFCell cell = row.getCell(cellSite);
        if(cell == null){
            cell = row.createCell(cellSite);
        }else if(!StringUtils.isEmpty(cell.toString()) && cell.toString().indexOf(this.start) < 0 ){
            if(isY){
                moveCellY(sheet, rowSite, cellSite, 1);
            }else{
                moveCellX(row, cellSite, 1);
            }
        }

        if(data instanceof Byte || data instanceof Short || data instanceof Integer ){
            cell.setCellValue((int) data);
        }else if(data instanceof Long ){
            cell.setCellValue((long) data);
        }else if(data instanceof Float || data instanceof Double){
            cell.setCellValue((double) data);
        }else if(data instanceof Boolean){
            cell.setCellValue((boolean) data);
        }else if(data instanceof Date){
            cell.setCellValue((Date) data);
        }else{
            cell.setCellValue(data.toString());
        }
        return cell;
    }

    public void eachTransferStop(String name, Object data, List<CellContent> cellContents,
                                 HSSFSheet sheet, int rowSite, int cellSite, int site){
        if(name.indexOf(this.x) >= 0){
            if(data instanceof Map) {
                setXMapData((HashMap) data, cellContents, sheet, site);
            }else if(data instanceof List){
                setXListData(name, (List<Object>) data, cellContents, sheet, rowSite, cellSite);
            }else{
                setXArrayData((Object[]) data, sheet, rowSite, cellSite);
            }
        }else {
            if (data instanceof Map) {
                setYMapData((HashMap) data, cellContents, sheet, site);
            } else if (data instanceof List) {
                setYListData(name, (List<Object>) data, cellContents, sheet, rowSite, cellSite);
            } else {
                setYArrayData((Object[]) data, sheet, rowSite, cellSite);
            }
        }
    }

    public void setXMapData(Map<String, Object> datas, List<CellContent> cellContents, HSSFSheet sheet, int site){
        int i = 0;
        for (Map.Entry<String, Object> data : datas.entrySet()){
            String key = data.getKey();
            Object value = data.getValue();
            for (CellContent cellContent : cellContents) {
                Integer rowSite = cellContent.getRowSite();
                Integer cellSite = cellContent.getCellSite() + (site > -9999 ? site : i);
                if(cellContent.getName().equals(key)){
                    if(!(value instanceof Map || value instanceof List || value.getClass().isArray())){
                        HSSFCell hssfCell = setCellValue(value, sheet, rowSite, cellSite);
                        setCellStyle(hssfCell, allCellStyle.get(cellContent.getName()));
                    }else{
                        eachTransferStop(key, value, cellContents, sheet, rowSite, cellSite, site);
                    }
                }
            }
        }
    }

    public void setYMapData(Map<String, Object> datas, List<CellContent> cellContents, HSSFSheet sheet, int site){
        int i = 0;
        for (Map.Entry<String, Object> data : datas.entrySet()){
            String key = data.getKey();
            Object value = data.getValue();
            for (CellContent cellContent : cellContents) {
                Integer rowSite = cellContent.getRowSite() + (site > -9999 ? site : i);
                Integer cellSite = cellContent.getCellSite();
                if(cellContent.getName().equals(key)){
                    if(!(value instanceof Map || value instanceof List || value.getClass().isArray())){
                        HSSFCell hssfCell = setCellValue(value, sheet, rowSite, cellSite);
                        setCellStyle(hssfCell, allCellStyle.get(cellContent.getName()));
                    }else{
                        eachTransferStop(key, value, cellContents, sheet, rowSite, cellSite, site);
                    }
                }
            }
        }
    }

    public void setXListData(String name, List<Object> datas, List<CellContent> cellContents, HSSFSheet sheet, int rowSite, int cellSite){
        for (int i = 0; i < datas.size(); i++) {
            Object o = datas.get(i);
            if(o instanceof Map || o instanceof List || o.getClass().isArray()){
                eachTransferStop(name, o, cellContents, sheet, rowSite, cellSite, i);
            }else {
                HSSFCell hssfCell = setCellValue(o, sheet, rowSite, cellSite + 1);
                setCellStyle(hssfCell, allCellStyle.get(name));
            }
        }
    }

    public void setYListData(String name, List<Object> datas, List<CellContent> cellContents, HSSFSheet sheet, int rowSite, int cellSite){
        for (int i = 0; i < datas.size(); i++) {
            Object o = datas.get(i);
            if(o instanceof Map || o instanceof List || o.getClass().isArray()){
                eachTransferStop(name, o, cellContents, sheet, rowSite, cellSite, i);
            }else {
                HSSFCell hssfCell = setCellValue(o, sheet, rowSite + i, cellSite, true);
                setCellStyle(hssfCell, allCellStyle.get(name));
            }
        }
    }

    public void setXArrayData(Object[] datas, HSSFSheet sheet, int rowSite, int cellSite){
        HSSFRow row = sheet.getRow(rowSite);
        HSSFCell cell = row.getCell(cellSite);
        for (int i = 0; i < datas.length; i++) {
            HSSFCell hssfCell = setCellValue(datas[i], sheet, rowSite, cellSite + 1);
            copyCellStyle(cell, hssfCell);
        }
    }

    public void setYArrayData(Object[] datas, HSSFSheet sheet, int rowSite, int cellSite){
        HSSFRow row = sheet.getRow(rowSite);
        HSSFCell cell = row.getCell(cellSite);
        for (int i = 0; i < datas.length; i++) {
            HSSFCell hssfCell = setCellValue(datas[i], sheet, rowSite + i, cellSite, true);
            copyCellStyle(cell, hssfCell);
        }
    }

    public void moveCellX(HSSFRow row, int cellSite,int time){
        HSSFCell cell = row.getCell(cellSite);
        for (int i = 0; i < time; i++) {
            HSSFCell newCell = row.getCell(cellSite + i + 1);
            if(newCell == null){
                newCell = row.createCell(cellSite + 1);
                CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                copyCellStyle(cell, newCell);
            }else if(StringUtils.isEmpty(newCell.toString())){
                CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                copyCellStyle(cell, newCell);
            }else{
                moveCellX(row, cellSite + i +1, time);
                CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
            }
        }
    }

    public void moveCellY(HSSFSheet sheet, int rowSite, int cellSite,int time){
        HSSFRow row = sheet.getRow(rowSite);
        HSSFCell cell = row.getCell(cellSite);
        for (int i = 0; i < time; i++) {
            HSSFRow newRow = sheet.getRow(rowSite + i +1);
            if(newRow == null){
                newRow = sheet.createRow(rowSite + 1);
                HSSFCell newCell = newRow.createCell(cellSite);
                CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                copyCellStyle(cell, newCell);
            }else{
                HSSFCell newCell = newRow.getCell(cellSite);
                if(newCell == null) {
                    newCell = row.createCell(cellSite);
                    CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                    copyCellStyle(cell, newCell);
                }else if(StringUtils.isEmpty(newCell.toString())){
                    CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                    copyCellStyle(cell, newCell);
                }else{
                    moveCellY(sheet, rowSite + i +1, cellSite, time);
                    CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                }
            }
        }
    }

    public void setCellStyle(HSSFCell cell, HSSFCellStyle cellStyle){
        if(cell != null && cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    public void copyCellStyle(HSSFCell cell, HSSFCell newCell){
        if(cell != null && newCell != null) {
            HSSFCellStyle cellStyle = cell.getCellStyle();
            if (cellStyle != null) {
                newCell.setCellStyle(cellStyle);
            }
        }
    }

}
