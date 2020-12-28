package org.example;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * 数据导出成excel
 * @author Administrator
 */
public class ExcelExport {
    private Map<String, CellInfo> cellInfo = new HashMap<>();
    private Map<String, Map<String, CellInfo>> arrayCellInfo = new HashMap<>();
    private Map<String, CellInfo> inUse = new HashMap<>();
    private Map<Integer, String[]> siteName = new HashMap<>();
    private Workbook wb = null;
    public static final String ORDER = "PRIORITY";
    private final String x = "X-";
    private String suffix = "";

    public void excelAdaptive(Workbook wb, Map<String, Object> dataMap) {
        this.wb = wb;
        excel(dataMap);
        adaptiveColumn(wb, 255);
    }

    public static Workbook excelAdaptive(String fileName, Map<String, Object> dataMap) throws IOException {
        ExcelExport excelExport = new ExcelExport();
        Workbook wb = excelExport.excel(fileName, dataMap);
        ExcelExport.adaptiveColumn(wb, 255);
        return wb;
    }

    public void excel(Map<String, Object> dataMap) {
        Sheet sheet = this.wb.getSheetAt(0);
        initialize(sheet, dataMap);
        setBasicData(sheet, dataMap);
        setAllArray(sheet, dataMap);
    }

    public Workbook excel(String fileName, Map<String, Object> dataMap) throws IOException {
        start(fileName);
        Sheet sheet = this.wb.getSheetAt(0);
        initialize(sheet, dataMap);
        setBasicData(sheet, dataMap);
        setAllArray(sheet, dataMap);
        return this.wb;
    }

    /**
     * 清空数据
     */
    public void empty(){
        this.cellInfo = new HashMap<>(0);
        this.arrayCellInfo = new HashMap<>(0);
        this.inUse = new HashMap<>(0);
        this.siteName = new HashMap<>(0);
        this.wb = null;
        this.suffix = "";
    }

    /**
     * 所有表自适应列宽（第1列开始）
     * @param wb excel文件
     * @param columnNum 列数
     */
    public static void adaptiveColumn(Workbook wb, int columnNum){
        adaptiveColumn(wb, 0, columnNum);
    }

    /**
     * 所有表自适应列宽
     * @param wb excel文件
     * @param firstCell 开始列
     * @param columnNum 列数
     */
    public static void adaptiveColumn(Workbook wb, int firstCell, int columnNum){
        int numberOfSheets = wb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = wb.getSheetAt(i);
            adaptiveColumn(sheet, firstCell, columnNum);
        }
    }

    /**
     * 自适应列宽
     * @param sheet 表
     * @param firstCell 开始列
     * @param columnNum 列数
     */
    public static void adaptiveColumn(Sheet sheet, int firstCell, int columnNum){
        for(int j = firstCell; j < columnNum; j++) {
            sheet.autoSizeColumn(j);
            sheet.setColumnWidth(j, (sheet.getColumnWidth(j) + 400));
        }
    }

    /**
     * 文件初始化
     * @param fileName 文件名
     * @throws IOException io流异常
     */
    public void start(String fileName) throws IOException {
        if(!fileName.endsWith(".xlsx")&&!fileName.endsWith(".xls")){
            throw new IOException("不是excel文件");
        }
        ClassPathResource resource = new ClassPathResource(fileName);
        this.suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        InputStream is = resource.getInputStream();
        this.wb = "xlsx".equals(this.suffix) ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }

    /**
     * 拷贝表
     * @param sheet 被拷贝表
     * @param newSheet 拷贝目地表
     */
    public static void copySheet(Sheet sheet, Sheet newSheet){
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        int maxColumnNum = 0;
        for (int i = firstRowNum; i <= lastRowNum; i++) {
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

//        if("xlsx".equals(this.suffix)){
//            XSSFSheet xSheet = (XSSFSheet) sheet;
//        }else{
//            HSSFSheet hSheet = (HSSFSheet) sheet;
//        }

    }

    /**
     * 拷贝行
     * @param row 被拷贝行
     * @param newRow 拷贝目地行
     */
    public static void copyRow(Row row, Row newRow){
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

    /**
     * 拷贝单元格
     * @param cell 被拷贝单元格
     * @param newCell 拷贝目地单元格
     */
    public static void copyCell(Cell cell, Cell newCell){
        if(cell == null || newCell == null){return;}
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
     * @param sheet 表
     * @param dataMap 数据
     */
    public void initialize(Sheet sheet, Map<String, Object> dataMap){
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
                        String start = "${";
                        String finish = "}";
                        int initially = cellValue.indexOf(start);
                        int ending = cellValue.indexOf(finish);
                        if (initially >= 0 && ending >= 0) {
                            String substring = cellValue.substring(initially + start.length(), ending);
                            String[] split = substring.split("\\.");
                            int splitLen = split.length;
                            String s = split[0];
                            Object o = dataMap.get(s);
                            if (o != null) {
                                String subEnding = cellValue.substring(ending + finish.length());
                                String str = split[splitLen - 1];
                                CellInfo cellInfo = new CellInfo(cellValue.substring(0, initially), str, subEnding, i, j);
                                cellInfo.setCellStyle(cell.getCellStyle());
                                cellInfo.setRowHeight(row.getHeight());
                                cellInfo.setColumnWidth(sheet.getColumnWidth(j));
                                cellInfo.setMergedResult(isMergedRegion(sheet, i, j));

                                if (isArray(o)) {
                                    Map<String, CellInfo> cellInfos = this.arrayCellInfo.get(s);
                                    if (CollectionUtils.isEmpty(cellInfos)) {
                                        cellInfos = new HashMap<>(5);
                                    }
                                    cellInfos.put(str, cellInfo);
                                    this.siteName.put((i * 100 + j), new String[]{s,str});
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
     * @param sheet 表
     * @param dataMap 数据
     */
    public void setBasicData(Sheet sheet, Map<String, Object> dataMap){
        cellInfo.forEach((key, value) -> {
            Object o = dataMap.get(key);
            Integer rowIndex = value.getRowIndex();
            Row row = sheet.getRow(rowIndex);
            Integer cellIndex = value.getCellIndex();
            Cell cell = row.getCell(cellIndex);
            if(o instanceof Image){
                Image image = (Image) o;
                byte[] bytes = image.getBytes();
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                boolean merged = mergedRegion.isMerged();
                int startX = 0;
                int startY = 0;
                int endX = 255;
                int endY = 255;
                int startCell = (merged ? mergedRegion.getFirstColumn() : cellIndex);
                int startRow = (merged ? mergedRegion.getFirstRow() : rowIndex);
                int endCell = (merged ? mergedRegion.getLastColumn() : cellIndex);
                int endRow = (merged ? mergedRegion.getLastRow() : rowIndex);
                if("xlsx".equals(this.suffix)){
                    addXssfImage((XSSFSheet) sheet, bytes, startX, startY, endX, endY, startCell, startRow, (endCell + 1), (endRow + 1));
                }else if("xls".equals(this.suffix)){
                    addHssfImage((HSSFSheet) sheet, bytes, startX, startY, (endX + 1023 - endX),  (endY - 10), (short) startCell, startRow, (short) endCell, endRow);
                }
            }else {
                String initially = value.getInitially();
                String ending = value.getEnding();
                cell.setCellValue(initially + o.toString() + ending);
            }
        });
    }

    /**
     * xlsx文件插入图片
     * @param sheet 表
     * @param bytes 图片二进制数组
     * @param startX 开始X坐标
     * @param startY 开始Y坐标
     * @param endX 结束X坐标
     * @param endY 结束Y坐标
     * @param startCell 开始列
     * @param startRow 开始行
     * @param endCell 结束列
     * @param endRow 结束行
     */
    public void addXssfImage(XSSFSheet sheet, byte[] bytes, int startX, int startY, int endX, int endY, int startCell, int startRow, int endCell, int endRow){
        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor(startX, startY, endX, endY, startCell, startRow, endCell, endRow);
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        int i = this.wb.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);
        drawingPatriarch.createPicture(anchor, i);
    }

    /**
     * xls文件插入图片
     * @param sheet 表
     * @param bytes 图片二进制数组
     * @param startX 开始X坐标
     * @param startY 开始Y坐标
     * @param endX 结束X坐标
     * @param endY 结束Y坐标
     * @param startCell 开始列
     * @param startRow 开始行
     * @param endCell 结束列
     * @param endRow 结束行
     */
    public void addHssfImage(HSSFSheet sheet, byte[] bytes, int startX, int startY, int endX, int endY, short startCell, int startRow, short endCell, int endRow){
        HSSFPatriarch drawingPatriarch = sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(startX, startY, endX, endY, startCell, startRow, endCell, endRow);
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        int i = this.wb.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);
        drawingPatriarch.createPicture(anchor, i);
    }

    /**
     * 所有遍历赋值单元格赋值
     * @param sheet 表
     * @param dataMap 数据
     */
    public void setAllArray(Sheet sheet, Map<String, Object> dataMap){
        String[] order = (String[]) dataMap.get(ORDER);
        if(order != null){
            for (String key : order) {
                Object o = dataMap.get(key);
                this.inUse = arrayCellInfo.get(key);
                if(this.inUse == null){
                    initialize(sheet, dataMap);
                    this.inUse = arrayCellInfo.get(key);
                }
                if(this.inUse != null){
                    CellInfo cellInfo = this.inUse.get(key);
                    if(cellInfo != null){
                        eachTransferStop(key, o, sheet, cellInfo.getRowIndex(), cellInfo.getCellIndex(), 1);
                    }else{
                        eachTransferStop(key, o, sheet);
                    }
                }
            }
        }else {
            arrayCellInfo.forEach((key, value) -> {
                Object o = dataMap.get(key);
                this.inUse = value;
                if(this.inUse != null){
                    CellInfo cellInfo = this.inUse.get(key);
                    if(cellInfo != null){
                        eachTransferStop(key, o, sheet, cellInfo.getRowIndex(), cellInfo.getCellIndex(), 1);
                    }else{
                        eachTransferStop(key, o, sheet);
                    }
                }

            });
        }
    }

    /**
     * 左右移动合并单元格赋值
     * @param data 要赋的值
     * @param mergedResult 合并单元格信息
     * @param sheet 表
     * @param rowIndex 行号
     * @param cellIndex 列号
     * @return 复制的单元格
     */
    public Cell setMergedCellValueX(Object data, MergedResult mergedResult, Sheet sheet, int rowIndex, int cellIndex){
        int firstColumn = mergedResult.getFirstColumn();
        int firstRow = mergedResult.getFirstRow();
        if(!(firstRow == rowIndex && firstColumn == cellIndex)){
            int rowNum = mergedResult.getRowMergeNum();
            int columnNum = mergedResult.getColumnMergeNum();
            int lastCell = cellIndex + columnNum - 1;
            moveRegionX(sheet, rowIndex, rowNum, cellIndex - 1, columnNum);
            List<List<CellStyle>> regionStyle = getRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);
            setRegionStyle(sheet, regionStyle, rowIndex, cellIndex);
            CellRangeAddress cra = new CellRangeAddress(rowIndex, rowIndex + rowNum - 1, cellIndex, lastCell);
            sheet.addMergedRegion(cra);
            List<Integer> cellWidths = getRegionCellWidth(sheet, firstColumn, columnNum);
            setRegionCellWidth(sheet, cellWidths, cellIndex);
        }
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(cellIndex);
        if(cell == null){
            cell = row.createCell(cellIndex);
        }
        setCellValue(cell, data);
        return cell;
    }

    /**
     * 左右移动单元给赋值
     * @param data 数据
     * @param sheet 表
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @return 赋值的单元格
     */
    public Cell setCellValueX(Object data, Sheet sheet, int rowIndex, int cellIndex, int moveNum){
        Row row = sheet.getRow(rowIndex);
        Cell cell ;
        if(row == null){
            row = sheet.createRow(rowIndex);
            cell = row.createCell(rowIndex);
        } else {
            cell = row.getCell(cellIndex);
            MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
            if (mergedRegion.isMerged()) {
                moveMergeCellX(sheet, mergedRegion, moveNum);
            } else if(cell != null && !StringUtils.isEmpty(cell.toString())){
                moveCellX(sheet, rowIndex, cellIndex, moveNum);
            }
            if (cell == null) {
                cell = row.createCell(cellIndex);
            }
        }
        setCellValue(cell, data);
        return cell;
    }

    public Cell setCellValueX(Object data, Sheet sheet, int rowIndex, int cellIndex){
        return setCellValueX(data, sheet, rowIndex, cellIndex, 1);
    }

    /**
     * 上下移动合并单元格赋值
     * @param data 要赋的值
     * @param mergedResult 合并单元格信息
     * @param sheet 表
     * @param rowIndex 行号
     * @param cellIndex 列号
     */
    public void setMergedCellValueY(Object data, MergedResult mergedResult, Sheet sheet, int rowIndex, int cellIndex){
        int firstRow = mergedResult.getFirstRow();
        int firstColumn = mergedResult.getFirstColumn();
        if(!(firstRow == rowIndex && firstColumn == cellIndex)){
            int rowNum = mergedResult.getRowMergeNum();
            int columnNum = mergedResult.getColumnMergeNum();
            int lastRow = rowIndex + rowNum - 1;
            moveRegionY(sheet, rowIndex - 1, rowNum, cellIndex, columnNum);
            List<List<CellStyle>> regionStyle = getRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);
            setRegionStyle(sheet, regionStyle, rowIndex, cellIndex);
            CellRangeAddress cra = new CellRangeAddress(rowIndex, lastRow, cellIndex, cellIndex + columnNum - 1);
            sheet.addMergedRegion(cra);
            List<Short> rowHeights = getRegionRowHeight(sheet, firstRow, rowNum);
            setRegionRowHeight(sheet, rowHeights, rowIndex);
        }
        Row row = sheet.getRow(rowIndex);
        if(row == null){
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(cellIndex);
        if(cell == null){
            cell = row.createCell(cellIndex);
        }
        setCellValue(cell, data);
    }

    /**
     * 上下单元给赋值
     * @param data 数据
     * @param sheet 表
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @return 赋值的单元格
     */
    public Cell setCellValueY(Object data, Sheet sheet, int rowIndex, int cellIndex, int moveNum){
        Row row = sheet.getRow(rowIndex);
        Cell cell ;
        if(row == null){
            row = sheet.createRow(rowIndex);
            cell = row.createCell(cellIndex);
        }else {
            cell = row.getCell(cellIndex);
            MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
            if(mergedRegion.isMerged()){
                moveMergeCellY(sheet, mergedRegion, moveNum);
            } else if(cell != null && !StringUtils.isEmpty(cell.toString())){
                moveCellY(sheet, rowIndex, cellIndex, moveNum);
            }
            if (cell == null) {
                cell = row.createCell(cellIndex);
            }
        }
        setCellValue(cell, data);
        return cell;
    }

    public Cell setCellValueY(Object data, Sheet sheet, int rowIndex, int cellIndex){
        return setCellValueY(data, sheet, rowIndex, cellIndex, 1);
    }

    public void setCellValue(Cell cell, Object data){
        if(data instanceof Number){
            cell.setCellValue(((Number) data).doubleValue());
            cell.setCellType(CellType.NUMERIC);
        }else if(data instanceof Boolean){
            cell.setCellValue((boolean) data);
            cell.setCellType(CellType.BOOLEAN);
        }else{
            cell.setCellValue(data.toString());
        }
    }

    /**
     * 中转站，根据数据类型调用相应方法进行赋值（Map，List，Object[]）
     * @param name 赋值单元格名称
     * @param data 赋值数据
     * @param sheet 表
     * @param rowIndex 初始行
     * @param cellIndex 初始列
     * @param index 初始列偏移量  >-9999初始列+index,<=-9999初始列根据数据数量进行累加
     * @return 条数
     */
    public int eachTransferStop(String name, Object data, Sheet sheet, int rowIndex, int cellIndex, int index){
        if(data != null) {
            int size;
            if (name.contains(this.x)) {
                if (data instanceof Map) {
                    Map<String, Object> dataMap = new HashMap<>(10);
                    ((Map<?, ?>) data).forEach((key, value) -> dataMap.put(key.toString(), value));
                    setMapDataX(dataMap, sheet, index);
                    size = dataMap.size();
                } else if (data instanceof List) {
                    List<Object> dataList = new ArrayList<>(((List<?>) data));
                    setListDataX(name, dataList, sheet, rowIndex, cellIndex);
                    size = dataList.size();
                } else {
                    setArrayDataX(name, (Object[]) data, sheet, rowIndex, cellIndex);
                    size = ((Object[]) data).length;
                }
            } else {
                if (data instanceof Map) {
                    Map<String, Object> dataMap = new HashMap<>(10);
                    ((Map<?, ?>) data).forEach((key, value) -> dataMap.put(key.toString(), value));
                    setMapDataY(dataMap, sheet, index);
                    size = dataMap.size();
                } else if (data instanceof List) {
                    List<Object> dataList = new ArrayList<>(((List<?>) data));
                    setListDataY(name, dataList, sheet, rowIndex, cellIndex);
                    size = dataList.size();
                } else {
                    setArrayDataY(name, (Object[]) data, sheet, rowIndex, cellIndex);
                    size = ((Object[]) data).length;
                }
            }
            return size;
        }
        return 0;
    }

    public void eachTransferStop(String name, Object data, Sheet sheet){
        eachTransferStop(name, data, sheet, 0, 0, -10000);
    }

    /**
     * Map集合向右赋值
     * @param dataMap 赋值数据
     * @param sheet 表
     * @param index 初始列偏移量  >-9999初始列+index,<=-9999初始列根据数据数量进行累加
     */
    public void setMapDataX(Map<String, Object> dataMap, Sheet sheet, int index){
        int i = 0;
        String[] priority = (String[]) dataMap.get(ORDER);
        if(priority != null){
            for (String key : priority) {
                Object data = dataMap.get(key);
                if(data != null) {
                    mapDataX(sheet, key, data, (index > -9999 ? index : i++));
                }
            }
        }else {
            for (Map.Entry<String, Object> data : dataMap.entrySet()) {
                mapDataX(sheet, data.getKey(), data.getValue(), (index > -9999 ? index : i++));
            }
        }
    }

    public void mapDataX(Sheet sheet, String key, Object value, int index){
        CellInfo cellInfo = this.inUse.get(key);
        if(cellInfo != null) {
            MergedResult mergedResult = cellInfo.getMergedResult();
            boolean merged = mergedResult.isMerged();
            int columnNum = mergedResult.getColumnMergeNum();
            Integer rowIndex = cellInfo.getRowIndex();
            Integer cellIndex = cellInfo.getCellIndex();
            int cellIndexI = cellIndex + index;
            if (isNoArray(value)) {
                if(index!=0){
                    moveCellSiteX(rowIndex, cellIndexI, 1);
                }
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if (!StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending)) {
                    value = initially + value + ending;
                }
                Cell cell;
                if(merged){
                    cell = setMergedCellValueX(value, mergedResult, sheet, rowIndex, cellIndexI);
                    cellInfo.setCellIndex(cellIndex + columnNum - 1);
                }else{
                    cell = setCellValueX(value, sheet, rowIndex, cellIndexI);
                }
                setCellStyle(cell, cellInfo.getCellStyle());
                sheet.setColumnWidth(cellIndexI, cellInfo.getColumnWidth());
            } else {
                int size = eachTransferStop(key, value, sheet, rowIndex, cellIndexI, index);
                if(key.contains(this.x)){
                    cellInfo.setCellIndex( cellIndex + size - 1);
                }
                if(merged){
                    int num = cellInfo.getCellIndex() + columnNum - 1;
                    cellInfo.setCellIndex(num);
                    mergedResult.setLastColumn(num + mergedResult.getColumnMergeNum());
                }
            }
        }
    }

    /**
     * Map集合向下赋值
     * @param dataMap 赋值数据
     * @param sheet 表
     * @param index 初始行偏移量  >-9999初始行+index,<=-9999根据数据数量进行累加
     */
    public void setMapDataY(Map<String, Object> dataMap, Sheet sheet, int index){
        int i = 0;
        String[] priority = (String[]) dataMap.get(ORDER);
        if(priority != null){
            for (String key : priority) {
                Object data = dataMap.get(key);
                mapDataY(sheet, key, data, (index > -9999 ? index : i));
                dataMap.remove(key);
            }
        }else {
            for (Map.Entry<String, Object> data : dataMap.entrySet()) {
                mapDataY(sheet, data.getKey(), data.getValue(), (index > -9999 ? index : i));
            }
        }
    }

    public void mapDataY(Sheet sheet, String key, Object value, int index){
        CellInfo cellInfo = this.inUse.get(key);
        if(cellInfo != null) {
            MergedResult mergedResult = cellInfo.getMergedResult();
            boolean merged = mergedResult.isMerged();
            int rowNum = mergedResult.getRowMergeNum();
            Integer rowIndex = cellInfo.getRowIndex();
            int rowIndexI = rowIndex + index;
            Integer cellIndex = cellInfo.getCellIndex();
            if (isNoArray(value)) {
                if(index!=0){
                    moveCellSiteY(sheet, rowIndexI, cellIndex, 1);
                }
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if (!StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending)) {
                    value = initially + value + ending;
                }
                if(merged){
                    setMergedCellValueY(value, mergedResult, sheet, rowIndexI, cellIndex);
                    cellInfo.setRowIndex(rowIndex + rowNum - 1);
                }else{
                    Cell cell = setCellValueY(value, sheet, rowIndexI, cellIndex);
                    setCellStyle(cell, cellInfo.getCellStyle());
                    Row row = sheet.getRow(rowIndexI);
                    row.setHeight(cellInfo.getRowHeight());
                }
            } else {
                int size = eachTransferStop(key, value, sheet, rowIndexI, cellIndex, index);
                if(!key.contains(this.x)){
                    cellInfo.setRowIndex( rowIndex + size - 1);
                }
                if(merged){
                    int num = cellInfo.getRowIndex() + rowNum - 1;
                    cellInfo.setRowIndex(num);
                    mergedResult.setLastRow(num + mergedResult.getRowMergeNum());
                }
            }
        }
    }

    /**
     * List集合向右赋值
     * @param name 赋值单元格名称
     * @param dataMap 赋值数据
     * @param sheet 表
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setListDataX(String name, List<Object> dataMap, Sheet sheet, int rowIndex, int cellIndex){
        CellInfo cellInfo = this.inUse.get(name);
        MergedResult mergedResult = null;
        boolean merged = false;
        String initially = null;
        String ending = null;
        if(cellInfo != null){
            mergedResult = cellInfo.getMergedResult();
            merged = mergedResult.isMerged();
            initially = cellInfo.getInitially();
            ending = cellInfo.getEnding();
        }
        boolean isAdTo = !StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending);
        int dataSize = dataMap.size();
        for (int i = 0; i < dataSize; i++) {
            Object value = dataMap.get(i);
            int cellIndexI = cellIndex + i;
            if(isNoArray(value)){
                if(i!=0){
                    moveCellSiteX(rowIndex, cellIndexI, 1);
                }
                if(isAdTo) {
                    value = initially + value + ending;
                }
                if(merged){
                    setMergedCellValueY(value, mergedResult, sheet, rowIndex, cellIndexI);
                    cellIndex += mergedResult.getColumnMergeNum() - 1;
                }else{
                    Cell cell = setCellValueX(value, sheet, rowIndex, cellIndexI, dataSize - i);
                    if (cellInfo != null) {
                        setCellStyle(cell, cellInfo.getCellStyle());
                        sheet.setColumnWidth(cellIndexI, cellInfo.getColumnWidth());
                    }
                }
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndexI, i);
                if(name.contains(this.x)){
                    cellIndex += size - 1;
                }
            }
        }
    }

    /**
     * List集合向下赋值
     * @param name 赋值单元格名称
     * @param dataMap 赋值数据
     * @param sheet 表
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setListDataY(String name, List<Object> dataMap, Sheet sheet, int rowIndex, int cellIndex){
        CellInfo cellInfo = this.inUse.get(name);
        MergedResult mergedResult = null;
        boolean merged = false;
        String initially = null;
        String ending = null;
        if(cellInfo != null){
            mergedResult = cellInfo.getMergedResult();
            merged = mergedResult.isMerged();
            initially = cellInfo.getInitially();
            ending = cellInfo.getEnding();
        }
        boolean isAddTo = !StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending);
        int dataSize = dataMap.size();
        for (int i = 0; i < dataSize; i++) {
            Object value = dataMap.get(i);
            int rowIndexI = rowIndex + i;
            if(isNoArray(value)){
                if(i!=0){
                    moveCellSiteY(sheet, rowIndexI, cellIndex, 1);
                }
                if(isAddTo) {
                    value = initially + value + ending;
                }
                if(merged){
                    setMergedCellValueY(value, mergedResult, sheet, rowIndexI, cellIndex);
                    rowIndex += mergedResult.getRowMergeNum() - 1;
                }else{
                    Cell cell = setCellValueY(value, sheet, rowIndexI, cellIndex, dataSize - i);
                    if (cellInfo != null) {
                        setCellStyle(cell, cellInfo.getCellStyle());
                        Row row = sheet.getRow(rowIndexI);
                        row.setHeight(cellInfo.getRowHeight());
                    }
                }
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndexI, cellIndex, i);
                if(!name.contains(this.x)){
                    rowIndex += size - 1;
                }
            }
        }
    }

    /**
     * 数组类型向右赋值
     * @param name 赋值单元格名称
     * @param dataMap 赋值数据
     * @param sheet 表
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setArrayDataX(String name, Object[] dataMap, Sheet sheet, int rowIndex, int cellIndex){
        CellInfo cellInfo = this.inUse.get(name);
        String initially = null;
        String ending = null;
        if(cellInfo != null){
            initially = cellInfo.getInitially();
            ending = cellInfo.getEnding();
        }
        boolean isAddTo = !StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending);
        int dataLen = dataMap.length;
        for (int i = 0; i < dataLen; i++, cellIndex++) {
            Object value = dataMap[i];
            if(isNoArray(value)){
                if(i!=0){
                    moveCellSiteX(rowIndex, cellIndex, 1);
                }
                if(isAddTo) {
                    value = initially + value + ending;
                }
                Cell cell = setCellValueX(value, sheet, rowIndex, cellIndex, dataLen - i);
                if (cellInfo != null) {
                    setCellStyle(cell, cellInfo.getCellStyle());
                    cellInfo.setColumnWidth(cellInfo.getColumnWidth());
                }
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(name.contains(this.x)){
                    cellIndex += size - 1;
                }
            }
        }
    }

    /**
     * 数组类型向下赋值
     * @param name 赋值单元格名称
     * @param dataMap 赋值数据
     * @param sheet 表
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setArrayDataY(String name, Object[] dataMap, Sheet sheet, int rowIndex, int cellIndex){
        CellInfo cellInfo = this.inUse.get(name);
        String initially = null;
        String ending = null;
        if(cellInfo != null){
            initially = cellInfo.getInitially();
            ending = cellInfo.getEnding();
        }
        boolean isAddTo = !StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending);
        int dataLen = dataMap.length;
        for (int i = 0; i < dataLen; i++, rowIndex++) {
            Object value = dataMap[i];
            if(isNoArray(value)){
                if(i!=0){
                    moveCellSiteX(rowIndex, cellIndex, 1);
                }
                if(isAddTo) {
                    value = initially + value + ending;
                }
                Cell cell = setCellValueY(value, sheet, rowIndex, cellIndex, dataLen - i);
                if (cellInfo != null) {
                    setCellStyle(cell, cellInfo.getCellStyle());
                    Row row = sheet.getRow(rowIndex);
                    row.setHeight(cellInfo.getRowHeight());
                }
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(!name.contains(this.x)){
                    rowIndex += size - 1;
                }
            }
        }
    }

    public void moveCellSiteX(int rowIndex, int cellIndex, int moveNum){
        int site = rowIndex * 100 + cellIndex;
        String[] name = this.siteName.get(site);
        if(!StringUtils.isEmpty(name)){
            Map<String, CellInfo> stringCellInfoMap = this.arrayCellInfo.get(name[0]);
            CellInfo cellInfo = stringCellInfoMap.get(name[1]);
            int cellSite = cellIndex + moveNum;
            cellInfo.setCellIndex(cellSite);
            this.siteName.remove(site);
            this.siteName.put((rowIndex * 100 + cellSite), name);
            MergedResult mergedResult = cellInfo.getMergedResult();
            if(mergedResult != null && mergedResult.isMerged()){
                mergedResult.setColumnIndex(cellSite);
                mergedResult.setFirstColumn(cellSite);
                mergedResult.setLastColumn(mergedResult.getLastColumn() + moveNum);
            }
        }
    }

    /**
     * 左右移动单元格
     * @param sheet 表
     * @param rowIndex 单元格所在行
     * @param cellIndex 单元格所在列
     * @param time 移动列数
     */
    public void moveCellX(Sheet sheet, int rowIndex, int cellIndex,int time){
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(cellIndex);
        Cell newCell = null;
        for (int i = 1; i <= time; i++) {
            int cIndex = cellIndex + i;
            newCell = row.getCell(cIndex);
            int moveNum = time - i + 1;
            moveCellSiteX(rowIndex, cIndex, moveNum);
            MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cIndex);
            if(mergedRegion.isMerged()){
                moveMergeCellX(sheet, mergedRegion, moveNum);
                newCell = row.getCell(cellIndex + time);
            }else if(newCell != null && !StringUtils.isEmpty(newCell.toString())){
                moveCellX(sheet, rowIndex, cIndex, moveNum);
                newCell = row.getCell(cellIndex + time);
                break;
            }else if(newCell == null){
                newCell = row.createCell(cIndex);
            }
        }
        copyCell(cell, newCell);
        copyCellStyle(cell, newCell);
        sheet.setColumnWidth(cellIndex + time, sheet.getColumnWidth(cellIndex));
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    public boolean moveCellSiteY(Sheet sheet, int rowIndex, int cellIndex, int moveNum){
        int site = rowIndex * 100 + cellIndex;
        String[] name = this.siteName.get(site);
        if(!StringUtils.isEmpty(name)){
            Map<String, CellInfo> stringCellInfoMap = this.arrayCellInfo.get(name[0]);
            CellInfo cellInfo = stringCellInfoMap.get(name[1]);
            int rowSite = rowIndex + moveNum;
            cellInfo.setRowIndex(rowSite);
            this.siteName.remove(site);
            this.siteName.put((rowSite * 100 + cellIndex), name);
            Row row = sheet.getRow(rowSite);
            if(row == null){
                row = sheet.createRow(rowSite);
            }
            row.setHeight(cellInfo.getRowHeight());
            MergedResult mergedResult = cellInfo.getMergedResult();
            if(mergedResult != null && mergedResult.isMerged()){
                mergedResult.setRowIndex(rowSite);
                mergedResult.setFirstRow(rowSite);
                mergedResult.setLastRow(mergedResult.getLastRow() + moveNum);
            }
            return true;
        }
        return false;
    }

    /**
     * 上下移动单元格
     * @param sheet 表
     * @param rowIndex 单元格所在行
     * @param cellIndex 单元格所在列
     * @param time 移动行数
     */
    public void moveCellY(Sheet sheet, int rowIndex, int cellIndex,int time){
        Row row = sheet.getRow(rowIndex);
        Row newRow = null;
        Cell cell = row.getCell(cellIndex);
        Cell newCell = null;
        for (int i = 1; i <= time ; i++) {
            int rIndex = rowIndex + i;
            newRow = sheet.getRow(rIndex);
            int moveNum = time - i + 1;
            moveCellSiteY(sheet, rIndex, cellIndex, moveNum);
            if(newRow == null){
                newRow = sheet.createRow(rIndex);
                newCell = newRow.createCell(cellIndex);
            }else{
                newCell = newRow.getCell(cellIndex);
                MergedResult mergedRegion = isMergedRegion(sheet, rIndex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellY(sheet, mergedRegion, moveNum);
                    newRow = sheet.getRow(rowIndex + time);
                    if(newRow == null){
                        newRow = sheet.createRow(rowIndex + time);
                    }
                    newCell = newRow.getCell(cellIndex);
                    break;
                }else if(newCell != null && !StringUtils.isEmpty(newCell.toString())){
                    moveCellY(sheet, rIndex, cellIndex, moveNum);
                    newRow = sheet.getRow(rowIndex + time);
                    newCell = newRow.getCell(cellIndex);
                    break;
                }else if(newCell == null){
                    newCell = newRow.createCell(cellIndex);
                }
            }
        }

        copyCell(cell, newCell);
        copyCellStyle(cell, newCell);
        if (newRow != null) {
            newRow.setHeight(row.getHeight());
        }
        if(cell != null) {
            cell.setCellStyle(null);
            cell.setCellValue("");
        }
    }

    /**
     * 左右移动合并单元格
     * @param sheet 表
     * @param mr 合并单元格信息
     * @param time 移动列数
     */
    public void moveMergeCellX(Sheet sheet, MergedResult mr,int time){
        int firstRow = mr.getFirstRow();
        int lastRow = mr.getLastRow();
        int firstColumn = mr.getFirstColumn();
        int lastColumn = mr.getLastColumn();
        int rowNum = mr.getRowMergeNum();
        int columnNum = mr.getColumnMergeNum();

        //获取合并单元格每列宽度
        List<Integer> columnWidths = getRegionCellWidth(sheet, firstColumn, columnNum);

        //判断移动到达的位置处有无值，将有值的向右移动
        moveRegionX(sheet, firstRow, rowNum, lastColumn, time);

        //拆分合并单元格
        splitMergedRegion(sheet, firstRow, firstColumn);

        //保存旧单元格的值
        Row row = sheet.getRow(firstRow);
        Cell cell = row.getCell(firstColumn);
        if(cell == null){
            cell = row.createCell(firstColumn);
        }
        String value = cell.toString();
        cell.setCellValue("");

        //保存合并单元格每个单元格的样式
        List<List<CellStyle>> regionStyle = getRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);

        //清除旧单元格样式
        clearRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);

        //将旧单元格的值设给新单元格
        int firstCellIndex = firstColumn + time;
        Cell newCell = row.getCell(firstCellIndex);
        if(newCell == null){
            newCell = row.createCell(firstCellIndex);
        }
        newCell.setCellValue(value);

        //给每个要合并的单元格设置样式
        setRegionStyle(sheet, regionStyle, firstRow, firstCellIndex);

        //为移动后的合并单元格每列设宽
        setRegionCellWidth(sheet, columnWidths, firstCellIndex);

        //合并单元格
        CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCellIndex, lastColumn + time);
        sheet.addMergedRegion(cra);

    }

    /**
     * 区域左右移动
     * @param sheet 表
     * @param firstRow 开始行
     * @param rowNum 移动行数
     * @param firstCell 开始列
     * @param cellNum 移动列数
     */
    public void moveRegionX(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        for (int i = 0; i < rowNum; i++) {
            int rowIndex = firstRow + i;
            Row row = sheet.getRow(rowIndex);
            for (int j = 1; j <= cellNum; j++) {
                int cellIndex = firstCell + j;
                Cell cell = row.getCell(cellIndex);
                int moveNum = cellNum - j + 1;
                boolean isEnd = false;
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellX(sheet, mergedRegion, moveNum);
                    isEnd = true;
                }else if(cell != null && !StringUtils.isEmpty(cell.toString())){
                    moveCellX(sheet, rowIndex, cellIndex, moveNum);
                }
                moveCellSiteX(rowIndex, cellIndex, moveNum);
                if(cell == null){
                    row.createCell(cellIndex);
                    if(isEnd){
                        break;
                    }
                }
            }
        }
    }

    /**
     * 清除区域内样式
     * @param sheet 表
     * @param firstRow 开始行数
     * @param rowNum 多少行
     * @param firstCell 开始列数
     * @param cellNum 多少列
     */
    public void clearRegionStyle(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        for (int i = 0; i < rowNum; i++) {
            Row row = sheet.getRow(firstRow + i);
            for (int j = 0; j < cellNum; j++) {
                Cell cell = row.getCell(firstCell + j);
                cell.setCellStyle(null);
            }
        }
    }

    /**
     * 获取区域样式
     * @param sheet 表
     * @param firstRow 开始行数
     * @param rowNum 多少行
     * @param firstCell 开始列数
     * @param cellNum 多少列
     * @return 区域内每个单元格的样式
     */
    public List<List<CellStyle>> getRegionStyle(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        List<List<CellStyle>> regionStyle = new ArrayList<>();
        for(int i = 0; i < rowNum; i++){
            Row row = sheet.getRow(firstRow + i);
            if(row == null){
                row = sheet.createRow(firstRow + i);
            }
            List<CellStyle> rowStyle = new ArrayList<>();
            for(int j = 0; j < cellNum; j++){
                Cell cell = row.getCell(firstCell + j);
                if(cell == null){
                    cell = row.createCell(firstCell + j);
                }
                rowStyle.add(cell.getCellStyle());
            }
            regionStyle.add(rowStyle);
        }
        return regionStyle;
    }

    /**
     * 设置区域样式
     * @param sheet 表
     * @param regionStyle 每个单元格的样式
     * @param firstRow 开始行
     * @param firstCell 开始列
     */
    public void setRegionStyle(Sheet sheet, List<List<CellStyle>> regionStyle, int firstRow, int firstCell){
        for (int i = 0; i < regionStyle.size(); i++) {
            List<CellStyle> styles = regionStyle.get(i);
            Row row = sheet.getRow(firstRow + i);
            if(row == null){
                row = sheet.createRow(firstRow + i);
            }
            for (int j = 0; j < styles.size(); j++) {
                Cell cell = row.getCell(firstCell + j);
                if(cell == null){
                    cell = row.createCell(firstCell + j);
                }
                cell.setCellStyle(styles.get(j));
            }
        }
    }

    /**
     * 获取区域行高
     * @param sheet 表
     * @param firstRow 开始行
     * @param rowNum 行数
     * @return 区域内每行的高
     */
    public List<Short> getRegionRowHeight(Sheet sheet, int firstRow, int rowNum){
        List<Short> rowHeights = new ArrayList<>();
        for (int i = 0; i < rowNum; i++) {
            int rowIndex = firstRow + i;
            Row row = sheet.getRow(rowIndex);
            if(row == null){
                row = sheet.createRow(rowIndex);
            }
            rowHeights.add(row.getHeight());
        }
        return rowHeights;
    }

    /**
     * 设置区域行高
     * @param sheet 表
     * @param rowHeights 要设置的行高
     * @param firstRow 开始行
     */
    public void setRegionRowHeight(Sheet sheet, List<Short> rowHeights, int firstRow){
        for (int i = 0; i < rowHeights.size(); i++) {
            Row row = sheet.getRow(firstRow + i);
            row.setHeight(rowHeights.get(i));
        }
    }

    /**
     * 获取区域列宽
     * @param sheet 表
     * @param firstColumn 开始列号
     * @param columnNum 列数
     * @return 区域内每列宽
     */
    public List<Integer> getRegionCellWidth(Sheet sheet, int firstColumn, int columnNum){
        List<Integer> columnWidths = new ArrayList<>();
        for (int i = 0; i < columnNum; i++) {
            columnWidths.add(sheet.getColumnWidth(firstColumn + i));
        }
        return columnWidths;
    }

    /**
     * 设置区域列宽
     * @param sheet 表
     * @param columnWidths 要设置的列宽
     * @param firstCellIndex 开始列号
     */
    public void setRegionCellWidth(Sheet sheet, List<Integer> columnWidths, int firstCellIndex) {
        for (int i = 0; i < columnWidths.size(); i++) {
            sheet.setColumnWidth(firstCellIndex + i, columnWidths.get(i));
        }
    }

    /**
     * 上下移动合并单元格
     * @param sheet 表
     * @param mr 合并单元格信息
     * @param time 移动行数
     */
    public void moveMergeCellY(Sheet sheet, MergedResult mr,int time){
        int firstRow = mr.getFirstRow();
        int lastRow = mr.getLastRow();
        int firstColumn = mr.getFirstColumn();
        int lastColumn = mr.getLastColumn();
        int rowNum = mr.getRowMergeNum();
        int columnNum = mr.getColumnMergeNum();

        List<Short> regionRowHeight = getRegionRowHeight(sheet, firstRow, rowNum);

        moveRegionY(sheet, lastRow, time, firstColumn, columnNum);

        splitMergedRegion(sheet, firstRow, firstColumn);

        Row row = sheet.getRow(firstRow);
        Cell cell = row.getCell(firstColumn);
        if(cell == null){
            cell = row.createCell(firstColumn);
        }
        String value = cell.toString();
        cell.setCellValue("");

        List<List<CellStyle>> regionStyle = getRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);

        clearRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);

        int firstRowIndex = firstRow + time;
        row = sheet.getRow(firstRowIndex);
        if(row == null){
            row = sheet.createRow(firstRowIndex);
        }
        Cell newCell = row.getCell(firstColumn);
        if (newCell == null) {
            newCell = row.createCell(firstColumn);
        }
        newCell.setCellValue(value);

        setRegionStyle(sheet, regionStyle, firstRowIndex, firstColumn);

        setRegionRowHeight(sheet, regionRowHeight, firstRowIndex);

        CellRangeAddress cra = new CellRangeAddress(firstRowIndex, lastRow + time, firstColumn, lastColumn);
        sheet.addMergedRegion(cra);

    }

    /**
     * 区域向下移动
     * @param sheet 表
     * @param firstRow 开始行号
     * @param rowNum 移动行数
     * @param firstCell 开始列号
     * @param cellNum 移动列数
     */
    public void moveRegionY(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        boolean[] isSkip = new boolean[cellNum + 1];
        int rowIndex = firstRow + 1;
        Row row = sheet.getRow(rowIndex);
        if(row != null) {
            for (int j = 0; j < cellNum; j++) {
                if (isSkip[j]) {
                    continue;
                }
                int cellIndex = firstCell + j;
                Cell cell = row.getCell(cellIndex);
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if (mergedRegion.isMerged()) {
                    moveMergeCellY(sheet, mergedRegion, rowNum);
                    for (int i1 = 0; i1 < mergedRegion.getColumnMergeNum(); i1++) {
                        isSkip[Math.min(isSkip.length - 1, j + i1)] = true;
                    }
                } else if (cell != null && !StringUtils.isEmpty(cell.toString())) {
                    moveCellY(sheet, rowIndex, cellIndex, rowNum);
                    isSkip[j] = true;
                } else if (cell == null) {
                    row.createCell(cellIndex);
                }
                boolean b = moveCellSiteY(sheet, rowIndex, cellIndex, rowNum);
                if (b) {
                    Row row1 = sheet.getRow(rowIndex + rowNum);
                    row1.setHeight(row.getHeight());
                }
            }
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
     * 判断对象是数组
     * 判断值是集合或者数组
     * @param o 要判断的对象
     * @return 是否是数组
     */
    public boolean isArray(Object o){
        return o instanceof Map || o instanceof List || o.getClass().isArray();
    }

    /**
     * 判断对象不是数组
     * 判断值不是集合或者数组
     * @param o 要判断的对象
     * @return 是否不是数组
     */
    public boolean isNoArray(Object o){
        return !isArray(o);
    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet 表
     * @param row 行下标
     * @param column 列下标
     * @return 合并单元格信息
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
     * 拆分区域内所有合并单元格
     * @param sheet 表
     * @param firstRow 开始行
     * @param lastRow 结束行
     * @param firstColumn 开始列
     * @param lastColumn 结束列
     */
    public static void splitMergedRegion(Sheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn){
        //获取合并单元格的数量
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            // 获取合并后的单元格
            CellRangeAddress range = sheet.getMergedRegion(i);
            if(range != null) {
                int stateRow = range.getFirstRow();
                int endRow = range.getLastRow();
                int stateColumn = range.getFirstColumn();
                int endColumn = range.getLastColumn();
                boolean inStateRow = stateRow >= firstRow && stateRow <= lastRow;
                boolean inEndRow = endRow >= firstRow && endRow <= lastRow;
                boolean inStateColumn = stateColumn >= firstColumn && stateColumn <= lastColumn;
                boolean inEndColumn = endColumn >= firstColumn && endColumn <= lastColumn;
                boolean topLeft = (inStateRow && inStateColumn);// 左上角
                boolean topRight = (inEndRow && inStateColumn);// 右上角
                boolean belowLeft = (inStateRow && inEndColumn);// 左下角
                boolean belowRight = (inEndRow && inEndColumn);// 右下角
                if (topLeft || topRight || belowLeft || belowRight) {//左下角 右下角
                    sheet.removeMergedRegion(i--);
                }
            }
        }
    }

    /**
     * 拆分合并单元格
     * @param sheet 表
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
