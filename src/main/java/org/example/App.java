package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.File;
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
    private Map<String, Map<String, CellInfo>> arrayCellInfo = new HashMap<>();
    private Map<String, CellInfo> inUse = null;
    private final String start = "${";
    private final String finish = "}";
    private final String x = "X-";

    public static void main(String[] args) throws IOException {
        ClassPathResource resource = new ClassPathResource("test.xlsx");
        InputStream is = resource.getInputStream();
//        POIFSFileSystem ps = new POIFSFileSystem(is);
//        Workbook wb = new HSSFWorkbook(is);
        Workbook wb = new XSSFWorkbook(is);
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
        File file = new File("D:/A临时/excel/test2.xlsx");
        file.createNewFile();
        FileOutputStream os = new FileOutputStream(file);
        wb.write(os);
        os.flush();
        os.close();
    }

    public void excel(Workbook wb, Map<String, Object> datas) {
        Sheet sheetAt = wb.getSheetAt(0);
        initialize(sheetAt, datas);
        setBasicData(sheetAt, datas);

        Sheet sheet = wb.createSheet();
        copySheet(sheetAt, sheet);

        setAllArray(sheet, datas);

    }

    public void copySheet(Sheet sheet, Sheet newSheet){
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        int maxColumnNum = 0;
        for (int i = firstRowNum; i < lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if(row != null) {
                Row newRow = newSheet.getRow(i);
                if (newRow == null) {
                    newRow = newSheet.createRow(i);
                }
                copyRow(row, newRow);
                short lastCellNum = row.getLastCellNum();
                if (lastCellNum > maxColumnNum) {
                    maxColumnNum = lastCellNum;
                }
            }
        }

        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            // 获取合并后的单元格
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
            newSheet.addMergedRegion(cra);
        }

        for (int i = 0; i < maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }

    }

    public void copyRow(Row row, Row newRow){
        short height = row.getHeight();
        newRow.setHeight(height);
        CellStyle rowStyle = row.getRowStyle();
        if (rowStyle != null) {
            newRow.setRowStyle(rowStyle);
        }
        float heightInPoints = row.getHeightInPoints();
        newRow.setHeightInPoints(heightInPoints);
        int rowNum = row.getRowNum();
        newRow.setRowNum(rowNum);
        boolean zeroHeight = row.getZeroHeight();
        newRow.setZeroHeight(zeroHeight);

        short firstCellNum = row.getFirstCellNum();
        short lastCellNum = row.getLastCellNum();

        for (int i = firstCellNum; i < lastCellNum; i++) {
            Cell cell = row.getCell(i);
            if(cell != null) {
                Cell newCell = newRow.getCell(i);
                if (newCell == null) {
                    newCell = newRow.createCell(i);
                }
                copyCell(cell, newCell);
            }
        }
    }

    public void copyCell(Cell cell, Cell newCell){
        CellType cellTypeEnum = cell.getCellTypeEnum();
        if(cellTypeEnum == CellType.STRING){
            newCell.setCellValue(cell.getStringCellValue());
        }else if (cellTypeEnum == CellType.NUMERIC){
            newCell.setCellValue(cell.getNumericCellValue());
        }else if (cellTypeEnum == CellType.BLANK){
            newCell.setCellType(cellTypeEnum);
        }else if (cellTypeEnum == CellType.BOOLEAN){
            newCell.setCellValue(cell.getBooleanCellValue());
        }else if (cellTypeEnum == CellType.ERROR){
            newCell.setCellErrorValue(cell.getErrorCellValue());
        }else if (cellTypeEnum == CellType.FORMULA){
            newCell.setCellFormula(cell.getCellFormula());
        }

        newCell.setCellStyle(cell.getCellStyle());
        newCell.setCellComment(cell.getCellComment());
        newCell.setHyperlink(cell.getHyperlink());
    }

    /**
     * 初始化，获取所有赋值单元格信息（${name}），数据分类（可直接赋值，遍历赋值（Map、List、Object[]））
     * @param sheet
     * @param datas 数据
     */
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
                                    Map<String, CellInfo> cellInfos = this.arrayCellInfo.get(s);
                                    if (CollectionUtils.isEmpty(cellInfos)) {
                                        cellInfos = new HashMap<>();
                                    }
                                    cellInfos.put(str, cellInfo);
                                    this.arrayCellInfo.put(s, cellInfos);
                                } else {
                                    this.cellInfo.put(s, cellInfo);
                                }

                                cell.setCellValue("");
                            }
                        }
                    }
                }
            }
        }

    }

    /**
     * 所有直接赋值单元格赋值
     * @param sheet
     * @param datas 数据
     */
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

    /**
     * 所有遍历赋值单元格赋值
     * @param sheet
     * @param datas 数据
     */
    public void setAllArray(Sheet sheet, Map<String, Object> datas){
        arrayCellInfo.forEach((key, value) -> {
            Object o = datas.get(key);
            this.inUse = value;
            eachTransferStop(key, o, sheet);
        });
    }

    /**
     * 给单元给赋值
     * @param data 数据
     * @param sheet
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @return
     */
    public Cell setCellValueX(Object data, Sheet sheet, int rowIndex, int cellIndex){
        Row row = sheet.getRow(rowIndex);
        if(row == null){
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(cellIndex);
        if(cell == null){
            cell = row.createCell(cellIndex);
        }else{
            String value = cell.toString();
            if(!StringUtils.isEmpty(value)){
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellX(sheet, mergedRegion, 1);
                }else {
                    moveCellX(sheet, rowIndex, cellIndex, 1);
                }
            }
        }
        setCellValue(cell, data);
        return cell;
    }

    /**
     * 给单元给赋值
     * @param data 数据
     * @param sheet
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @return
     */
    public Cell setCellValueY(Object data, Sheet sheet, int rowIndex, int cellIndex){
        Row row = sheet.getRow(rowIndex);
        if(row == null){
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(cellIndex);
        if(cell == null){
            cell = row.createCell(cellIndex);
        }else{
            String value = cell.toString();
            if(!StringUtils.isEmpty(value)){
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellY(sheet, mergedRegion, 1);
                }else {
                    moveCellY(sheet, rowIndex, cellIndex, 1);
                }
            }
        }
        setCellValue(cell, data);
        return cell;
    }

    public void setCellValue(Cell cell, Object data){
        if(data instanceof Number){
            cell.setCellValue(((Number) data).doubleValue());
        }else if(data instanceof Boolean){
            cell.setCellValue((boolean) data);
        }else if(data instanceof Date){
            cell.setCellValue((Date) data);
        }else{
            cell.setCellValue(data.toString());
        }
    }

    public int eachTransferStop(String name, Object data, Sheet sheet){
        return eachTransferStop(name, data, sheet, 0, 0, -10000);
    }

    /**
     * 中转站，根据数据类型调用相应方法进行赋值（Map，List，Object[]）
     * @param name 赋值单元格名称
     * @param data 赋值数据
     * @param sheet
     * @param rowIndex 初始行
     * @param cellIndex 初始列
     * @param index 初始列偏移量  >-9999初始列+index,<=-9999初始列根据数据数量进行累加
     * @return
     */
    public int eachTransferStop(String name, Object data, Sheet sheet, int rowIndex, int cellIndex, int index){
        int size;
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

    /**
     * Map集合向右赋值
     * @param datas 赋值数据
     * @param sheet
     * @param index 初始列偏移量  >-9999初始列+index,<=-9999初始列根据数据数量进行累加
     */
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
                    Cell Cell = setCellValueX(value, sheet, rowIndex, cellIndex);
                    setCellStyle(Cell, cellInfo.getCellStyle());
                    sheet.setColumnWidth(cellIndex, cellInfo.getColumnWidth());
                } else {
                    setCellValueX("", sheet, rowIndex, cellIndex);
                    int size = eachTransferStop(key, value, sheet, rowIndex, cellIndex, index);
                    if(key.indexOf(this.x) >= 0){
                        cellInfo.setCellIndex( cellInfo.getCellIndex() + size - 1);
                    }
                }
            }
        }
    }

    /**
     * Map集合向下赋值
     * @param datas 赋值数据
     * @param sheet
     * @param index 初始行偏移量  >-9999初始行+index,<=-9999根据数据数量进行累加
     */
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
                    Cell Cell = setCellValueY(value, sheet, rowIndex, cellIndex);
                    setCellStyle(Cell, cellInfo.getCellStyle());
                    Row row = sheet.getRow(rowIndex);
                    row.setHeight(cellInfo.getRowHeigth());
                } else {
                    setCellValueY("", sheet, rowIndex, cellIndex);
                    int size = eachTransferStop(key, value, sheet, rowIndex, cellIndex, index);
                    if(!(key.indexOf(this.x) >= 0)){
                        cellInfo.setRowIndex( cellInfo.getRowIndex() + size - 1);
                    }
                }
            }
        }
    }

    /**
     * List集合向右赋值
     * @param name 赋值单元格名称
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
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
                Cell Cell = setCellValueX(value, sheet, rowIndex, cellIndex + i);
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

    /**
     * List集合向下赋值
     * @param name 赋值单元格名称
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
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
                Cell Cell = setCellValueY(value, sheet, rowIndex + i, cellIndex);
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

    /**
     * 数组类型向右赋值
     * @param name 赋值单元格名称
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
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
                Cell Cell = setCellValueX(value, sheet, rowIndex, cellIndex + i);
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

    /**
     * 数组类型向下赋值
     * @param name 赋值单元格名称
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
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
                Cell Cell = setCellValueY(value, sheet, rowIndex + i, cellIndex);
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

    /**
     * 判断单元格类型（普通单元格，合并单元格）调用相关方法进行移动
     * 如果单元格是赋值单元格（${name}）,更新单元格信息CellInfo
     * @param sheet
     * @param value 单元格值
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @param moveNum 移动列数
     */
    public void moveX(Sheet sheet, String value, int rowIndex, int cellIndex, int moveNum){
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
            Map<String, CellInfo> cellInfos = this.arrayCellInfo.get(split[0]);
            CellInfo cellInfo = cellInfos.get(split[splitLen - 1]);
            cellInfo.setCellIndex(cellInfo.getCellIndex() + moveNum);
        }
    }

    /**
     * 左右移动单元格
     * @param sheet
     * @param rowIndex 单元格所在行
     * @param cellIndex 单元格所在列
     * @param time 移动列数
     */
    public void moveCellX(Sheet sheet, int rowIndex, int cellIndex,int time){
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(cellIndex);
        Cell newCell = null;
        boolean isPlus = time > 0;
        for (int i = (isPlus ? 1 : -1); (isPlus ? i <= time : i >= time); i+= (isPlus ? 1 : -1)) {
            int cindex = cellIndex + i;
            newCell = row.getCell(cindex);
            int moveNum = time - i + 1;
            if(newCell != null ){
                String value = newCell.toString();
                if(!StringUtils.isEmpty(value)){
                    moveX(sheet, value, rowIndex, cindex, moveNum);
                    break;
                }
            }else {
                if(newCell == null){
                    MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                    if(mergedRegion.isMerged()){
                        moveMergeCellY(sheet, mergedRegion, moveNum);
                    }
                    newCell = row.createCell(cellIndex + 1);
                }
            }
        }
        copyCell(cell, newCell);
        copyCellStyle(cell, newCell);
        sheet.setColumnWidth(cellIndex + time, sheet.getColumnWidth(cellIndex));
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    /**
     * 判断单元格类型（普通单元格，合并单元格）调用相关方法进行移动
     * 如果单元格是赋值单元格（${name}）,更新单元格信息CellInfo
     * @param sheet
     * @param value 单元格值
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @param moveNum 移动列数
     */
    public void moveY(Sheet sheet, String value, int rowIndex, int cellIndex, int moveNum){
        MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
        if (mergedRegion.isMerged()) {
            moveMergeCellY(sheet, mergedRegion, moveNum);
        } else {
            moveCellY(sheet, rowIndex, cellIndex, moveNum);
        }
        if(value.indexOf(this.start) >= 0 && value.indexOf(this.finish) >= 0){
            int initially = value.indexOf(this.start);
            int ending = value.indexOf(this.finish);
            String substring = value.substring(initially + this.start.length(), ending);
            String[] split = substring.split("\\.");
            int splitLen = split.length;
            Map<String, CellInfo> cellInfos = this.arrayCellInfo.get(split[0]);
            CellInfo cellInfo = cellInfos.get(split[splitLen - 1]);
            cellInfo.setRowIndex(cellInfo.getRowIndex() + moveNum);
        }
}

    /**
     * 上下移动单元格
     * @param sheet
     * @param rowIndex 单元格所在行
     * @param cellIndex 单元格所在列
     * @param time 移动行数
     */
    public void moveCellY(Sheet sheet, int rowIndex, int cellIndex,int time){
        Row row = sheet.getRow(rowIndex);
        Row newRow = null;
        Cell cell = row.getCell(cellIndex);
        Cell newCell = null;
        boolean isPlus = time > 0;
        for (int i = (isPlus ? 1 : -1); (isPlus ? i <= time : i >= time); i+= (isPlus ? 1 : -1)) {
            int rindex = rowIndex + i;
            newRow = sheet.getRow(rindex);
            if(newRow == null){
                newRow = sheet.createRow(rindex);
                newCell = newRow.createCell(cellIndex);
            }else{
                newCell = newRow.getCell(cellIndex);
                int moveNum = time - i + 1;
                if(newCell == null) {
                    MergedResult mergedRegion = isMergedRegion(sheet, newRow.getRowNum(), cellIndex);
                    if(mergedRegion.isMerged()){
                        moveMergeCellY(sheet, mergedRegion, moveNum);
                    }
                    newCell = newRow.createCell(cellIndex);
                }else{
                    String value = newCell.toString();
                    if(!StringUtils.isEmpty(value)){
                        moveY(sheet, value, rindex, cellIndex, moveNum);
                        break;
                    }
                }
            }
        }

        copyCell(cell, newCell);
        copyCellStyle(cell, newCell);
        newRow.setHeight(row.getHeight());
        cell.setCellStyle(null);
        cell.setCellValue("");
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

        //获取保存合并单元格每列宽度
        for (int i = 0; i < columnNum; i++) {
            sheet.getColumnWidth(firstColumn + i);
        }

        //判断移动到达的位置处有无值，将有值的向右移动
        for (int i = 0; i < rowNum; i++) {
            int rowIndex = firstRow + i;
            Row row = sheet.getRow(rowIndex);
            for (int j = 1; j < columnNum + time; j++) {
                int cellIndex = lastColumn + j;
                Cell cell = row.getCell(cellIndex);
                if (cell != null) {
                    String value = cell.toString();
                    if (!StringUtils.isEmpty(value)) {
                        int moveNum = time - j + 1;
                        moveX(sheet, value, rowIndex, cellIndex, moveNum);
                        break;
                    }
                }
            }
        }

        splitMergedRegion(sheet, firstRow, firstColumn);

        //保存旧单元格的值及样式
        Row row = sheet.getRow(firstRow);
        Cell cell = row.getCell(firstColumn);
        String value = cell.toString();
        CellStyle cellStyle = cell.getCellStyle();
        cell.setCellStyle(null);
        cell.setCellValue("");

        //将旧单元格的值及样式设给新单元格
        int firstCellIndex = firstColumn + time;
        Cell newCell = row.getCell(firstCellIndex);
        if(newCell == null){
            newCell = row.createCell(firstCellIndex);
        }
        newCell.setCellValue(value);
        newCell.setCellStyle(cellStyle);

        //为移动后的合并单元格每列蛇宽
        for (int i = 0; i < columnWidths.size(); i++) {
            sheet.setColumnWidth(firstCellIndex + i, columnWidths.get(i));
        }

        //合并单元格
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
        List<Short> rowWidths = new ArrayList<>();
        int firstRow = mr.getFirstRow();
        int lastRow = mr.getLastRow();
        int firstColumn = mr.getFirstColumn();
        int lastColumn = mr.getLastColumn();
        int rowNum = mr.getRowMergeNum();
        int columnNum = mr.getColumnMergeNum();
        for (int i = 0; i < rowNum; i++) {
            Row row = sheet.getRow(firstRow + i);
            if(row != null) {
                short height = row.getHeight();
                rowWidths.add(height);
            }
        }

        for (int i = 1; i < rowNum + time; i++) {
            int rowIndex = lastRow + i;
            Row row = sheet.getRow(rowIndex);
            if(row == null){
                sheet.createRow(rowIndex);
                continue;
            }
            for (int j = 0; j < columnNum; j++) {
                int cellIndex = firstColumn + j;
                Cell cell = row.getCell(cellIndex);
                if (cell != null) {
                    String value = cell.toString();
                    if (!StringUtils.isEmpty(value)) {
                        int moveNum = time - j + 1;
                        moveX(sheet, value, rowIndex, cellIndex, moveNum);
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
        cell.setCellStyle(null);
        cell.setCellValue("");

        int firstRowIndex = firstRow + time;
        row = sheet.getRow(firstRowIndex);
        if(row != null) {
            Cell newCell = row.getCell(firstColumn);
            if (newCell == null) {
                newCell = row.createCell(firstColumn);
            }
            newCell.setCellValue(value);
            newCell.setCellStyle(cellStyle);
        }
        for (int i = 0; i < rowWidths.size(); i++) {
            row = sheet.getRow(firstRowIndex + i);
            if(row != null) {
                row.setHeight(rowWidths.get(i));
            }
        }

        CellRangeAddress cra = new CellRangeAddress(firstRowIndex, lastRow + time, firstColumn, lastColumn);
        sheet.addMergedRegion(cra);

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
