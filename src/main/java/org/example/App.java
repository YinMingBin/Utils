package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
    private Map<String, CellInfo> cellInfo = new HashMap<>();
    private Map<String, Map<String, CellInfo>> ArrayCellInfo = new HashMap<>();
    private Map<String, CellInfo> inUse = null;
    private final String start = "${";
    private final String finish = "}";
    private final String x = "X-";

    public static void main(String[] args) throws IOException {
//        ClassPathResource resource = new ClassPathResource("test.xls");
        InputStream is = new FileInputStream("target/classes/test.xls");
        POIFSFileSystem ps = new POIFSFileSystem(is);
        Workbook wb = new HSSFWorkbook(ps);
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
            paramRecordMap.put("no", no);
            paramRecordMap.put("no2", no);
            paramRecords.add(paramRecordMap);
        }

        datas.put("X-paramRecords", paramRecords);

        App app = new App();
        app.excel(wb, datas);
        FileOutputStream os = new FileOutputStream("D:/A临时/excel/test2.xls");
        wb.write(os);
        os.flush();
        os.close();
    }

    public void excel(Workbook wb, Map<String, Object> datas) {
        Sheet sheetAt = wb.getSheetAt(0);
        initialize(sheetAt, datas);
        setBasicData(sheetAt, datas);

        Sheet sheet = wb.createSheet();
        CopySheetUtil.copySheets(sheet, sheetAt);

        setAllArray(sheet, datas);

    }

    public void initialize(Sheet sheet, Map<String, Object> datas){
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for(int i = firstRowNum; i <= lastRowNum; i++){
            Row row = sheet.getRow(i);
            if(row != null) {
                short firstCellNum = row.getFirstCellNum();
                short lastCellNum = row.getLastCellNum();
                for (int j = firstCellNum; j < lastCellNum; j++) {
                    Cell cell = row.getCell(j);
                    if(cell != null) {
                        String cellValue = cell.toString();
                        int initially = cellValue.indexOf(this.start);
                        int ending = cellValue.indexOf(this.finish);
                        if (initially >= 0 && ending >= 0) {
                            String substring = cellValue.substring(initially + this.start.length(), ending);
                            String[] split = substring.split("\\.");
                            int splitLen = split.length;
                            String s = split[0];
                            Object o = datas.get(s);
                            if (o != null) {
                                String subEnding = cellValue.substring(ending + this.finish.length());
                                String str = split[splitLen - 1];
                                CellInfo cellInfo = new CellInfo(cellValue.substring(0, initially), str, subEnding, i, j);
                                cellInfo.setCellStyle(cell.getCellStyle());
                                cellInfo.setRowHeigth(row.getHeight());
                                cellInfo.setColumnWidth(sheet.getColumnWidth(j));

                                if (isArray(o)) {
                                    Map<String, CellInfo> cellInfos = this.ArrayCellInfo.get(s);
                                    if (CollectionUtils.isEmpty(cellInfos)) {
                                        cellInfos = new HashMap<>();
                                    }
                                    cellInfos.put(str, cellInfo);
                                    this.ArrayCellInfo.put(s, cellInfos);
                                } else {
                                    this.cellInfo.put(s, cellInfo);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public void setBasicData(Sheet sheet, Map<String, Object> datas){
        cellInfo.forEach((key, value) -> {
            Object o = datas.get(key);
            Row row = sheet.getRow(value.getRowIndex());
            Cell cell = row.getCell(value.getCellIndex());
            String initially = value.getInitially();
            String ending = value.getEnding();
            cell.setCellValue(initially + o.toString() + ending);
        });
    }

    public void setAllArray(Sheet sheet, Map<String, Object> datas){
        ArrayCellInfo.forEach((key, value) -> {
            Object o = datas.get(key);
            this.inUse = value;
            eachTransferStop(key, o, sheet);
        });
    }

    public Cell setCellValue(String name, Object data, Sheet sheet, int rowIndex, int cellIndex){
        return setCellValue(name, data, sheet, rowIndex, cellIndex, false, false);
    }

    public Cell setCellValue(String name, Object data, Sheet sheet, int rowIndex, int cellIndex, boolean ifMove){
        return setCellValue(name, data, sheet, rowIndex, cellIndex, ifMove, false);
    }

    public Cell setCellValue(String name, Object data, Sheet sheet, int rowIndex, int cellIndex, boolean ifMove, boolean isY){
        Row row = sheet.getRow(rowIndex);
        if(row == null){
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(cellIndex);
        if(cell == null){
            cell = row.createCell(cellIndex);
        }else{
            String value = cell.toString();
            if(ifMove && !StringUtils.isEmpty(value) && value.indexOf(this.start) < 0 && value.indexOf(this.finish) < 0){
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if(isY){
                    if(mergedRegion.isMerged()){
                        moveMergeCellY(sheet, mergedRegion, 1);
                    }else {
                        moveCellY(sheet, rowIndex, cellIndex, 1);
                    }
                }else{
                    if(mergedRegion.isMerged()){
                        moveMergeCellX(sheet, mergedRegion, 1);
                    }else {
                        moveCellX(sheet, rowIndex, cellIndex, 1);
                    }
                }
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

    public int eachTransferStop(String name, Object data, Sheet sheet){
        return eachTransferStop(name, data, sheet, 0, 0, -10000);
    }

    public int eachTransferStop(String name, Object data, Sheet sheet, int rowIndex, int cellIndex, int index){
        int size = 0;
        if(name.indexOf(this.x) >= 0){
            if(data instanceof Map) {
                setXMapData((HashMap) data, sheet, index);
                size = ((HashMap) data).size();
            }else if(data instanceof List){
                setXListData(name, (List<Object>) data, sheet, rowIndex, cellIndex);
                size = ((List<Object>) data).size();
            }else{
                setXArrayData(name, (Object[]) data, sheet, rowIndex, cellIndex);
                size = ((Object[]) data).length;
            }
        }else {
            if (data instanceof Map) {
                setYMapData((HashMap) data, sheet, index);
                size = ((HashMap) data).size();
            } else if (data instanceof List) {
                setYListData(name, (List<Object>) data, sheet, rowIndex, cellIndex);
                size = ((List<Object>) data).size();
            } else {
                setYArrayData(name, (Object[]) data, sheet, rowIndex, cellIndex);
                size = ((Object[]) data).length;
            }
        }
        return size;
    }

    public void setXMapData(Map<String, Object> datas, Sheet sheet, int index){
        int i = 0;
        for (Map.Entry<String, Object> data : datas.entrySet()){
            String key = data.getKey();
            Object value = data.getValue();
            CellInfo cellInfo = this.inUse.get(key);
            if(cellInfo != null) {
                Integer rowIndex = cellInfo.getRowIndex();
                Integer cellIndex = cellInfo.getCellIndex() + (index > -9999 ? index : i);
                if (isNoArray(value)) {
                    String initially = cellInfo.getInitially();
                    String ending = cellInfo.getEnding();
                    if (!(StringUtils.isEmpty(initially) && StringUtils.isEmpty(ending))) {
                        value = initially + value + ending;
                    }
                    Cell Cell = setCellValue(key, value, sheet, rowIndex, cellIndex, (i > 0 || index > 0));
                    setCellStyle(Cell, cellInfo.getCellStyle());
                    sheet.setColumnWidth(cellIndex, cellInfo.getColumnWidth());
                } else {
                    int size = eachTransferStop(key, value, sheet, rowIndex, cellIndex, index);
                    if(key.indexOf(this.x) >= 0){
                        cellInfo.setCellIndex( cellInfo.getCellIndex() + size - 1);
                    }
                }
            }
        }
    }

    public void setYMapData(Map<String, Object> datas, Sheet sheet, int index){
        int i = 0;
        for (Map.Entry<String, Object> data : datas.entrySet()){
            String key = data.getKey();
            Object value = data.getValue();
            CellInfo cellInfo = this.inUse.get(key);
            if(cellInfo != null) {
                Integer rowIndex = cellInfo.getRowIndex() + (index > -9999 ? index : i);
                Integer cellIndex = cellInfo.getCellIndex();
                if (isNoArray(value)) {
                    String initially = cellInfo.getInitially();
                    String ending = cellInfo.getEnding();
                    if (!(StringUtils.isEmpty(initially) && StringUtils.isEmpty(ending))) {
                        value = initially + value + ending;
                    }
                    Cell Cell = setCellValue(key, value, sheet, rowIndex, cellIndex, (i > 0 || index > 0));
                    setCellStyle(Cell, cellInfo.getCellStyle());
                    Row row = sheet.getRow(rowIndex);
                    row.setHeight(cellInfo.getRowHeigth());
                } else {
                    int size = eachTransferStop(key, value, sheet, rowIndex, cellIndex, index);
                    if(!(key.indexOf(this.x) >= 0)){
                        cellInfo.setRowIndex( cellInfo.getRowIndex() + size - 1);
                    }
                }
            }
        }
    }

    public void setXListData(String name, List<Object> datas, Sheet sheet, int rowIndex, int cellIndex){
        for (int i = 0; i < datas.size(); i++) {
            Object value = datas.get(i);
            if(isNoArray(value)){
                CellInfo cellInfo = this.inUse.get(name);
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if(!(StringUtils.isEmpty(initially) && StringUtils.isEmpty(ending))) {
                    value = initially + value + ending;
                }
                Cell Cell = setCellValue(name, value, sheet, rowIndex, cellIndex + i, (i > 0));
                setCellStyle(Cell, cellInfo.getCellStyle());
                sheet.setColumnWidth(cellIndex + 1, cellInfo.getColumnWidth());
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(name.indexOf(this.x) >= 0){
                    cellIndex += size - 1;
                }
            }
        }
    }

    public void setYListData(String name, List<Object> datas, Sheet sheet, int rowIndex, int cellIndex){
        for (int i = 0; i < datas.size(); i++) {
            Object value = datas.get(i);
            if(isNoArray(value)){
                CellInfo cellInfo = this.inUse.get(name);
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if(!(StringUtils.isEmpty(initially) && StringUtils.isEmpty(ending))) {
                    value = initially + value + ending;
                }
                Cell Cell = setCellValue(name, value, sheet, rowIndex + i, cellIndex, (i > 0), true);
                setCellStyle(Cell, cellInfo.getCellStyle());
                Row row = sheet.getRow(rowIndex + i);
                row.setHeight(cellInfo.getRowHeigth());
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(!(name.indexOf(this.x) >= 0)){
                    rowIndex += size - 1;
                }
            }
        }
    }

    public void setXArrayData(String name, Object[] datas, Sheet sheet, int rowIndex, int cellIndex){
        for (int i = 0; i < datas.length; i++) {
            Object value = datas[i];
            if(isNoArray(value)){
                CellInfo cellInfo = this.inUse.get(name);
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if(!(StringUtils.isEmpty(initially) && StringUtils.isEmpty(ending))) {
                    value = initially + value + ending;
                }
                Cell Cell = setCellValue(name, value, sheet, rowIndex, cellIndex + i, (i > 0));
                setCellStyle(Cell, cellInfo.getCellStyle());
                cellInfo.setColumnWidth(cellInfo.getColumnWidth());
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(name.indexOf(this.x) >= 0){
                    cellIndex += size - 1;
                }
            }
        }
    }

    public void setYArrayData(String name, Object[] datas, Sheet sheet, int rowIndex, int cellIndex){
        for (int i = 0; i < datas.length; i++) {
            Object value = datas[i];
            if(isNoArray(value)){
                CellInfo cellInfo = this.inUse.get(name);
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if(!(StringUtils.isEmpty(initially) && StringUtils.isEmpty(ending))) {
                    value = initially + value + ending;
                }
                Cell Cell = setCellValue(name, value, sheet, rowIndex + i, cellIndex, (i > 0), true);
                setCellStyle(Cell, cellInfo.getCellStyle());
                Row row = sheet.getRow(rowIndex + i);
                row.setHeight(cellInfo.getRowHeigth());
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(!(name.indexOf(this.x) >= 0)){
                    rowIndex += size - 1;
                }
            }
        }
    }

    public void move(Sheet sheet, String value, int rowIndex, int cellIndex, int moveNum){
        MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
        if (mergedRegion.isMerged()) {
            moveMergeCellX(sheet, mergedRegion, moveNum);
        } else {
            moveCellX(sheet, rowIndex, cellIndex, moveNum);
        }
        if(value.indexOf(this.start) >= 0 && value.indexOf(this.finish) >= 0){
            int initially = value.indexOf(this.start);
            int ending = value.indexOf(this.finish);
            String substring = value.substring(initially + this.start.length(), ending);
            String[] split = substring.split("\\.");
            int splitLen = split.length;
            Map<String, CellInfo> cellInfos = this.ArrayCellInfo.get(split[0]);
            CellInfo cellInfo = cellInfos.get(split[splitLen - 1]);
            cellInfo.setCellIndex(cellInfo.getCellIndex() + moveNum);
        }
    }

    public void moveCellX(Sheet sheet, int rowIndex, int cellIndex,int time){
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(cellIndex);
        Cell newCell = null;
        for (int i = 1; i <= time; i++) {
            int cindex = cellIndex + i;
            newCell = row.getCell(cindex);
            if(newCell != null ){
                String value = newCell.toString();
                if(!StringUtils.isEmpty(value)){
                    int moveNum = time - i + 1;
                    move(sheet, value, rowIndex, cindex, moveNum);
                    break;
                }
            }else {
                if(newCell == null){
                    newCell = row.createCell(cellIndex + 1);
                }
            }
        }
        CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
        copyCellStyle(cell, newCell);
        sheet.setColumnWidth(cellIndex + time, sheet.getColumnWidth(cellIndex));
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    public void moveCellY(Sheet sheet, int rowIndex, int cellIndex,int time){
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(cellIndex);
        for (int i = 0; i < time; i++) {
            Row newRow = sheet.getRow(rowIndex + i + 1);
            if(newRow == null){
                newRow = sheet.createRow(rowIndex + 1);
                Cell newCell = newRow.createCell(cellIndex);
                CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                copyCellStyle(cell, newCell);
            }else{
                Cell newCell = newRow.getCell(cellIndex);
                if(newCell == null) {
                    newCell = row.createCell(cellIndex);
                    CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                    copyCellStyle(cell, newCell);
                }else if(StringUtils.isEmpty(newCell.toString())){
                    CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                    copyCellStyle(cell, newCell);
                }else{
                    moveCellY(sheet, rowIndex + i + 1, cellIndex, time);
                    CopySheetUtil.copyCell(cell, newCell, new HashMap<>());
                }
            }
        }
    }

    /**
     * 左右移动合并单元格
     * @param sheet
     * @param mr 合并单元格信息
     * @param time 移动列数
     */
    public void moveMergeCellX(Sheet sheet, MergedResult mr,int time){
        List<Integer> columnWidths = new ArrayList<>();
        int firstRow = mr.getFirstRow();
        int lastRow = mr.getLastRow();
        int firstColumn = mr.getFirstColumn();
        int lastColumn = mr.getLastColumn();
        int rowNum = mr.getRowMergeNum();
        int columnNum = mr.getColumnMergeNum();
        for (int i = 0; i < rowNum; i++) {
            int rowIndex = firstRow + i;
            Row row = sheet.getRow(rowIndex);
            for (int j = 1; j < columnNum + time; j++) {
                int cellIndex = lastColumn + j;
                if (i == 0) {
                    columnWidths.add(sheet.getColumnWidth(cellIndex));
                }
                Cell cell = row.getCell(cellIndex);
                if (cell != null) {
                    String value = cell.toString();
                    if (!StringUtils.isEmpty(value)) {
                        int moveNum = time - j + 1;
                        move(sheet, value, rowIndex, cellIndex, moveNum);
                        break;
                    }
                }
            }
        }

        splitMergedRegion(sheet, firstRow, firstColumn);

        Row row = sheet.getRow(firstRow);
        Cell cell = row.getCell(firstColumn);
        String value = cell.toString();
        CellStyle cellStyle = cell.getCellStyle();
        cell.setCellValue("");

        int firstCellIndex = firstColumn + time;
        Cell newCell = row.getCell(firstCellIndex);
        if(newCell == null){
            newCell = row.createCell(firstCellIndex);
        }
        newCell.setCellValue(value);
        newCell.setCellStyle(cellStyle);

        for (int i = 0; i < columnWidths.size(); i++) {
            sheet.setColumnWidth(firstCellIndex + i, columnWidths.get(i));
        }

        CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCellIndex, lastColumn + time);
        sheet.addMergedRegion(cra);

    }

    /**
     * 上下移动合并单元格
     * @param sheet
     * @param mr 合并单元格信息
     * @param time 移动行数
     */
    public void moveMergeCellY(Sheet sheet, MergedResult mr,int time){
        List<CellInfo> cellInfos = new ArrayList<>();
        int rowIndex = mr.getRowIndex();
        int columnIndex = mr.getColumnIndex();
        int rowNum = mr.getRowMergeNum();
        int columnNum = mr.getColumnMergeNum();
        for (int i = 0; i < (rowNum > columnNum ? rowNum : columnNum); i++) {
            Row row = sheet.getRow(rowIndex + (i < rowNum ? i : 0));
            CellInfo cellInfo = new CellInfo();
            cellInfo.setRowHeigth(row.getHeight());
            cellInfo.setColumnWidth(sheet.getColumnWidth(columnIndex + (i < columnNum ? i : 0)));
            cellInfos.add(cellInfo);
        }
    }

    /**
     * 设置单元格样式
     * @param cell 目标单元格
     * @param cellStyle 单元格样式
     */
    public void setCellStyle(Cell cell, CellStyle cellStyle){
        if(cell != null && cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 赋值单元格样式
     * @param cell 被复制的单元格
     * @param newCell 目标单元格
     */
    public void copyCellStyle(Cell cell, Cell newCell){
        if(cell != null && newCell != null) {
            CellStyle cellStyle = cell.getCellStyle();
            if (cellStyle != null) {
                newCell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 判断值是集合或者数组
     * @param o
     * @return
     */
    public boolean isArray(Object o){
        return o instanceof Map || o instanceof List || o.getClass().isArray();
    }

    /**
     * 判断值不是集合或者数组
     * @param o
     * @return
     */
    public boolean isNoArray(Object o){
        return !isArray(o);
    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row 行下标
     * @param column 列下标
     * @return
     */
    public static MergedResult isMergedRegion(Sheet sheet, int row, int column) {
        MergedResult mergedResult = new MergedResult();
        boolean isMerged = false;//判断是否合并单元格

        mergedResult.setRowIndex(row);//判断的行
        mergedResult.setColumnIndex(column);//判断的列
        //获取sheet中有多少个合并单元格
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            // 获取合并后的单元格
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow //判断行
                    && column >= firstColumn && column <= lastColumn) {//判断列
                isMerged = true;
                mergedResult.setFirstRow(firstRow);
                mergedResult.setLastRow(lastRow);
                mergedResult.setFirstColumn(firstColumn);
                mergedResult.setLastColumn(lastColumn);
                mergedResult.setRowMergeNum(lastRow - firstRow + 1);
                mergedResult.setColumnMergeNum(lastColumn - firstColumn + 1);
                break;
            }
        }
        mergedResult.setMerged(isMerged);
        return mergedResult;
    }

    /**
     * 拆分合并单元格
     * @param sheet
     * @param row 所在行
     * @param column 所在列
     */
    public void splitMergedRegion(Sheet sheet, int row, int column){
        //获取合并单元格的数量
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            // 获取合并后的单元格
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow //判断行
                    && column >= firstColumn && column <= lastColumn) {//判断列
                sheet.removeMergedRegion(i);
                break;
            }
        }
    }
}
