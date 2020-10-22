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

public class ExcelUtil {
    private Map<String, CellInfo> cellInfo = new HashMap<>();
    private Map<String, Map<String, CellInfo>> arrayCellInfo = new HashMap<>();
    private Map<String, CellInfo> inUse = new HashMap<>();
    private Map<Integer, String[]> siteName = new HashMap<>();
    private Workbook wb = null;
    public static final String ORDER = "PRIORITY";
    private final String x = "X-";
    private String suffix = "";

    public void empty(){
        this.cellInfo = new HashMap<>();
        this.arrayCellInfo = new HashMap<>();
        this.inUse = new HashMap<>();
        this.siteName = new HashMap<>();
        this.wb = null;
        this.suffix = "";
    }

    public void excelAdaptive(Workbook wb, Map<String, Object> datas) {
        this.wb = wb;
        excel(datas);
        adaptiveColumn(wb, 255);
    }

    public static Workbook excelAdaptive(String fileName, Map<String, Object> datas) throws IOException {
        ExcelUtil abc = new ExcelUtil();
        Workbook wb = abc.excel(fileName, datas);
        abc.adaptiveColumn(wb, 255);
        return wb;
    }

    public void adaptiveColumn(Workbook wb, int columnNum){
        adaptiveColumn(wb, 0, columnNum);
    }

    public void adaptiveColumn(Workbook wb, int firstCell, int columnNum){
        int numberOfSheets = wb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = wb.getSheetAt(i);
            adaptiveColumn(sheet, firstCell, columnNum);
        }
    }

    public void adaptiveColumn(Sheet sheet, int firstCell, int columnNum){
        for(int j = firstCell; j < columnNum; j++) {
            sheet.autoSizeColumn(j);
        }
    }

    public void start(String fileName) throws IOException {
        if(!fileName.endsWith(".xlsx")&&!fileName.endsWith(".xls")){
            throw new IOException("不是excel文件");
        }
        ClassPathResource resource = new ClassPathResource(fileName);
        this.suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        InputStream is = resource.getInputStream();
        this.wb = "xlsx".equals(this.suffix) ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }

    public Workbook pageExcel(String fileName, Map<String, Object> datas) throws IOException {
        start(fileName);

        Sheet sheetAt = this.wb.getSheetAt(0);

        initialize(sheetAt, datas);
        setBasicData(sheetAt, datas);

        Sheet sheet = this.wb.createSheet();
        copySheet(sheetAt, sheet);

        setAllArray(sheet, datas);

        return null;
    }

    public void excel(Map<String, Object> datas) {
        Sheet sheet = this.wb.getSheetAt(0);
        initialize(sheet, datas);
        setBasicData(sheet, datas);
        setAllArray(sheet, datas);
    }

    public Workbook excel(String fileName, Map<String, Object> datas) throws IOException {
        start(fileName);
        Sheet sheet = this.wb.getSheetAt(0);
        initialize(sheet, datas);
        setBasicData(sheet, datas);
        setAllArray(sheet, datas);
        return this.wb;
    }

    public void copySheet(Sheet sheet, Sheet newSheet){
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
                        String start = "${";
                        String finish = "}";
                        int initially = cellValue.indexOf(start);
                        int ending = cellValue.indexOf(finish);
                        if (initially >= 0 && ending >= 0) {
                            String substring = cellValue.substring(initially + start.length(), ending);
                            String[] split = substring.split("\\.");
                            int splitLen = split.length;
                            String s = split[0];
                            Object o = datas.get(s);
                            if (o != null) {
                                String subEnding = cellValue.substring(ending + finish.length());
                                String str = split[splitLen - 1];
                                CellInfo cellInfo = new CellInfo(cellValue.substring(0, initially), str, subEnding, i, j);
                                cellInfo.setCellStyle(cell.getCellStyle());
                                cellInfo.setRowHeigth(row.getHeight());
                                cellInfo.setColumnWidth(sheet.getColumnWidth(j));
                                cellInfo.setMergedResult(isMergedRegion(sheet, i, j));

                                if (isArray(o)) {
                                    Map<String, CellInfo> cellInfos = this.arrayCellInfo.get(s);
                                    if (CollectionUtils.isEmpty(cellInfos)) {
                                        cellInfos = new HashMap<>();
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
     * @param sheet
     * @param datas 数据
     */
    public void setBasicData(Sheet sheet, Map<String, Object> datas){
        cellInfo.forEach((key, value) -> {
            Object o = datas.get(key);
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
                    addXSSFImage((XSSFSheet) sheet, bytes, startX, startY, endX, endY, startCell, startRow, (endCell + 1), (endRow + 1));
                }else if("xls".equals(this.suffix)){
                    addHSSFImage((HSSFSheet) sheet, bytes, startX, startY, (endX + 1023 - endX), endY, (short) startCell, startRow, (short) endCell, endRow);
                }
            }else {
                String initially = value.getInitially();
                String ending = value.getEnding();
                cell.setCellValue(initially + o.toString() + ending);
            }
        });
    }

    public void addXSSFImage(XSSFSheet sheet, byte[] bytes, int startX, int startY, int endX, int endY, int startCell, int startRow, int endCell, int endRow){
        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor(startX, startY, endX, endY, startCell, startRow, endCell, endRow);
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        int i = this.wb.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);
        drawingPatriarch.createPicture(anchor, i);
    }

    public void addHSSFImage(HSSFSheet sheet, byte[] bytes, int startX, int startY, int endX, int endY, short startCell, int startRow, short endCell, int endRow){
        HSSFPatriarch drawingPatriarch = sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(startX, startY, endX, endY, startCell, startRow, endCell, endRow);
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        int i = this.wb.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);
        drawingPatriarch.createPicture(anchor, i);
    }

    /**
     * 所有遍历赋值单元格赋值
     * @param sheet
     * @param datas 数据
     */
    public void setAllArray(Sheet sheet, Map<String, Object> datas){
        String[] order = (String[]) datas.get(ORDER);
        if(order != null){
            for (String key : order) {
                Object o = datas.get(key);
                this.inUse = arrayCellInfo.get(key);
                if(this.inUse == null){
                    initialize(sheet, datas);
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
                Object o = datas.get(key);
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
     * 给单元给赋值
     * @param data 数据
     * @param sheet
     * @param rowIndex 所在行
     * @param cellIndex 所在列
     * @return
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

    public Cell setMergedCellValueY(Object data, MergedResult mergedResult, Sheet sheet, int rowIndex, int cellIndex){
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
        if(name.contains(this.x)){
            if(data instanceof Map) {
                setXMapData((Map) data, sheet, index);
                size = ((Map) data).size();
            }else if(data instanceof List){
                setXListData(name, (List<Object>) data, sheet, rowIndex, cellIndex);
                size = ((List<Object>) data).size();
            }else{
                setXArrayData(name, (Object[]) data, sheet, rowIndex, cellIndex);
                size = ((Object[]) data).length;
            }
        }else {
            if (data instanceof Map) {
                setYMapData((Map) data, sheet, index);
                size = ((Map) data).size();
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
        String[] priority = (String[]) datas.get(ORDER);
        if(priority != null){
            for (String key : priority) {
                Object data = datas.get(key);
                if(data != null) {
                    mapDataX(sheet, key, data, (index > -9999 ? index : i++));
                }
            }
        }else {
            for (Map.Entry<String, Object> data : datas.entrySet()) {
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
            Integer cellIndexI = cellIndex + index;
            if (isNoArray(value)) {
                if(index!=0){
                    moveXCellSite(rowIndex, cellIndexI, 1);
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
     * @param datas 赋值数据
     * @param sheet
     * @param index 初始行偏移量  >-9999初始行+index,<=-9999根据数据数量进行累加
     */
    public void setYMapData(Map<String, Object> datas, Sheet sheet, int index){
        int i = 0;
        String[] priority = (String[]) datas.get(ORDER);
        if(priority != null){
            for (String key : priority) {
                Object data = datas.get(key);
                mapDataY(sheet, key, data, (index > -9999 ? index : i));
                datas.remove(key);
            }
        }else {
            for (Map.Entry<String, Object> data : datas.entrySet()) {
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
            Integer rowIndexI = rowIndex + index;
            Integer cellIndex = cellInfo.getCellIndex();
            if (isNoArray(value)) {
                if(index!=0){
                    moveYCellSite(rowIndexI, cellIndex, 1);
                }
                String initially = cellInfo.getInitially();
                String ending = cellInfo.getEnding();
                if (!StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending)) {
                    value = initially + value + ending;
                }
                Cell Cell;
                if(merged){
                    Cell = setMergedCellValueY(value, mergedResult, sheet, rowIndexI, cellIndex);
                    cellInfo.setRowIndex(rowIndex + rowNum - 1);
                }else{
                    Cell = setCellValueY(value, sheet, rowIndexI, cellIndex);
                }
                setCellStyle(Cell, cellInfo.getCellStyle());
                Row row = sheet.getRow(rowIndexI);
                row.setHeight(cellInfo.getRowHeigth());
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
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setXListData(String name, List<Object> datas, Sheet sheet, int rowIndex, int cellIndex){
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
        int dataSize = datas.size();
        for (int i = 0; i < dataSize; i++) {
            Object value = datas.get(i);
            int cellIndexI = cellIndex + i;
            if(isNoArray(value)){
                if(i!=0){
                    moveXCellSite(rowIndex, cellIndexI, 1);
                }
                if(isAdTo) {
                    value = initially + value + ending;
                }
                if(merged){
                    setMergedCellValueY(value, mergedResult, sheet, rowIndex, cellIndexI);
                    cellIndex += mergedResult.getColumnMergeNum() - 1;
                }else{
                    Cell cell = setCellValueX(value, sheet, rowIndex, cellIndexI, dataSize - i);
                    setCellStyle(cell, cellInfo.getCellStyle());
                    sheet.setColumnWidth(cellIndexI, cellInfo.getColumnWidth());
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
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setYListData(String name, List<Object> datas, Sheet sheet, int rowIndex, int cellIndex){
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
        int dataSize = datas.size();
        for (int i = 0; i < dataSize; i++) {
            Object value = datas.get(i);
            int rowIndexI = rowIndex + i;
            if(isNoArray(value)){
                if(i!=0){
                    moveXCellSite(rowIndexI, cellIndex, 1);
                }
                if(isAddTo) {
                    value = initially + value + ending;
                }
                if(merged){
                    setMergedCellValueY(value, mergedResult, sheet, rowIndexI, cellIndex);
                    rowIndex += mergedResult.getRowMergeNum() - 1;
                }else{
                    Cell cell = setCellValueY(value, sheet, rowIndexI, cellIndex, dataSize - i);
                    setCellStyle(cell, cellInfo.getCellStyle());
                    Row row = sheet.getRow(rowIndexI);
                    row.setHeight(cellInfo.getRowHeigth());
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
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setXArrayData(String name, Object[] datas, Sheet sheet, int rowIndex, int cellIndex){
        CellInfo cellInfo = this.inUse.get(name);
        String initially = null;
        String ending = null;
        if(cellInfo != null){
            initially = cellInfo.getInitially();
            ending = cellInfo.getEnding();
        }
        boolean isAddTo = !StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending);
        int dataLen = datas.length;
        for (int i = 0; i < dataLen; i++, cellIndex++) {
            Object value = datas[i];
            if(isNoArray(value)){
                if(i!=0){
                    moveXCellSite(rowIndex, cellIndex, 1);
                }
                if(isAddTo) {
                    value = initially + value + ending;
                }
                Cell Cell = setCellValueX(value, sheet, rowIndex, cellIndex, dataLen - i);
                setCellStyle(Cell, cellInfo.getCellStyle());
                cellInfo.setColumnWidth(cellInfo.getColumnWidth());
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
     * @param datas 赋值数据
     * @param sheet
     * @param rowIndex 初始所在行
     * @param cellIndex 初始所在列
     */
    public void setYArrayData(String name, Object[] datas, Sheet sheet, int rowIndex, int cellIndex){
        CellInfo cellInfo = this.inUse.get(name);
        String initially = null;
        String ending = null;
        if(cellInfo != null){
            initially = cellInfo.getInitially();
            ending = cellInfo.getEnding();
        }
        boolean isAddTo = !StringUtils.isEmpty(initially) || !StringUtils.isEmpty(ending);
        int dataLen = datas.length;
        for (int i = 0; i < dataLen; i++, rowIndex++) {
            Object value = datas[i];
            if(isNoArray(value)){
                if(i!=0){
                    moveXCellSite(rowIndex, cellIndex, 1);
                }
                if(isAddTo) {
                    value = initially + value + ending;
                }
                Cell Cell = setCellValueY(value, sheet, rowIndex, cellIndex, dataLen - i);
                setCellStyle(Cell, cellInfo.getCellStyle());
                Row row = sheet.getRow(rowIndex);
                row.setHeight(cellInfo.getRowHeigth());
            }else {
                int size = eachTransferStop(name, value, sheet, rowIndex, cellIndex, i);
                if(!name.contains(this.x)){
                    rowIndex += size - 1;
                }
            }
        }
    }

    public boolean moveXCellSite(int rowIndex, int cellIndex, int moveNum){
        int site = rowIndex * 100 + cellIndex;
        String[] name = this.siteName.get(site);
        if(!StringUtils.isEmpty(name)){
            Map<String, CellInfo> stringCellInfoMap = this.arrayCellInfo.get(name[0]);
            CellInfo cellInfo = stringCellInfoMap.get(name[1]);
            int cellSite = cellIndex + moveNum;
            cellInfo.setCellIndex(cellSite);
            this.siteName.remove(site);
            this.siteName.put((rowIndex * 100 + cellSite), name);
            return true;
        }
        return false;
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
        for (int i = 1; i <= time; i++) {
            int cindex = cellIndex + i;
            newCell = row.getCell(cindex);
            int moveNum = time - i + 1;
            moveXCellSite(rowIndex, cindex, moveNum);
            MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cindex);
            if(mergedRegion.isMerged()){
                moveMergeCellX(sheet, mergedRegion, moveNum);
                newCell = row.getCell(cellIndex + time);
            }else if(newCell != null && !StringUtils.isEmpty(newCell.toString())){
                moveCellX(sheet, rowIndex, cindex, moveNum);
                newCell = row.getCell(cellIndex + time);
                break;
            }else if(newCell == null){
                newCell = row.createCell(cindex);
            }
        }
        copyCell(cell, newCell);
        copyCellStyle(cell, newCell);
        sheet.setColumnWidth(cellIndex + time, sheet.getColumnWidth(cellIndex));
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    public boolean moveYCellSite(int rowIndex, int cellIndex,int moveNum){
        int site = rowIndex * 100 + cellIndex;
        String[] name = this.siteName.get(site);
        if(!StringUtils.isEmpty(name)){
            Map<String, CellInfo> stringCellInfoMap = this.arrayCellInfo.get(name[0]);
            CellInfo cellInfo = stringCellInfoMap.get(name[1]);
            int rowSite = rowIndex + moveNum;
            cellInfo.setRowIndex(rowSite);
            this.siteName.remove(site);
            this.siteName.put((rowSite * 100 + cellIndex), name);
            return true;
        }
        return false;
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
        for (int i = 1; i <= time ; i++) {
            int rindex = rowIndex + i;
            newRow = sheet.getRow(rindex);
            if(newRow == null){
                newRow = sheet.createRow(rindex);
                newCell = newRow.createCell(cellIndex);
            }else{
                newCell = newRow.getCell(cellIndex);
                int moveNum = time - i + 1;
                moveYCellSite(rindex, cellIndex, moveNum);
                MergedResult mergedRegion = isMergedRegion(sheet, rindex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellY(sheet, mergedRegion, moveNum);
                    newRow = sheet.getRow(rowIndex + time);
                    newCell = newRow.getCell(cellIndex);
                    break;
                }else if(newCell != null && !StringUtils.isEmpty(newCell.toString())){
                    moveCellY(sheet, rindex, cellIndex, moveNum);
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

    public void moveRegionX(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        for (int i = 0; i < rowNum; i++) {
            int rowIndex = firstRow + i;
            Row row = sheet.getRow(rowIndex);
            for (int j = 1; j <= cellNum; j++) {
                int cellIndex = firstCell + j;
                Cell cell = row.getCell(cellIndex);
                int moveNum = cellNum - j + 1;
                boolean isEnd = false;
                moveXCellSite(rowIndex, cellIndex, moveNum);
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellX(sheet, mergedRegion, moveNum);
                    isEnd = true;
                }else if(cell != null && !StringUtils.isEmpty(cell.toString())){
                    moveCellX(sheet, rowIndex, cellIndex, moveNum);
                    break;
                }
                if(cell == null){
                    row.createCell(cellIndex);
                    if(isEnd){
                        break;
                    }
                }
            }
        }
    }

    public void clearRegionStyle(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        for (int i = 0; i < rowNum; i++) {
            Row row = sheet.getRow(firstRow + i);
            for (int j = 0; j < cellNum; j++) {
                Cell cell = row.getCell(firstCell + j);
                cell.setCellStyle(null);
            }
        }
    }

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

    public void setRegionRowHeight(Sheet sheet, List<Short> rowHeights, int firstRow){
        for (int i = 0; i < rowHeights.size(); i++) {
            Row row = sheet.getRow(firstRow + i);
            row.setHeight(rowHeights.get(i));
        }
    }

    public List<Integer> getRegionCellWidth(Sheet sheet, int firstColumn, int columnNum){
        List<Integer> columnWidths = new ArrayList<>();
        for (int i = 0; i < columnNum; i++) {
            columnWidths.add(sheet.getColumnWidth(firstColumn + i));
        }
        return columnWidths;
    }

    public void setRegionCellWidth(Sheet sheet, List<Integer> columnWidths, int firstCellIndex) {
        for (int i = 0; i < columnWidths.size(); i++) {
            sheet.setColumnWidth(firstCellIndex + i, columnWidths.get(i));
        }
    }

    /**
     * 上下移动合并单元格
     * @param sheet
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
        String value = cell.toString();
        cell.setCellValue("");

        List<List<CellStyle>> regionStyle = getRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);

        clearRegionStyle(sheet, firstRow, rowNum, firstColumn, columnNum);

        int firstRowIndex = firstRow + time;
        row = sheet.getRow(firstRowIndex);
        if(row != null) {
            Cell newCell = row.getCell(firstColumn);
            if (newCell == null) {
                newCell = row.createCell(firstColumn);
            }
            newCell.setCellValue(value);
        }

        setRegionStyle(sheet, regionStyle, firstRowIndex, firstColumn);

        setRegionRowHeight(sheet, regionRowHeight, firstRowIndex);

        CellRangeAddress cra = new CellRangeAddress(firstRowIndex, lastRow + time, firstColumn, lastColumn);
        sheet.addMergedRegion(cra);

    }

    public void moveRegionY(Sheet sheet, int firstRow, int rowNum, int firstCell, int cellNum){
        boolean[] isSkip = new boolean[rowNum + 1];
        for(int i = 1; i <= rowNum; i++){
            int rowIndex = firstRow + i;
            Row row = sheet.getRow(rowIndex);
            if(row == null){
                sheet.createRow(rowIndex);
                continue;
            }
            for (int j = 0; j < cellNum; j++) {
                if(isSkip[j]){ continue;}
                int cellIndex = firstCell + j;
                Cell cell = row.getCell(cellIndex);
                int moveNum = rowNum - i + 1;
                moveYCellSite(rowIndex, cellIndex, moveNum);
                MergedResult mergedRegion = isMergedRegion(sheet, rowIndex, cellIndex);
                if(mergedRegion.isMerged()){
                    moveMergeCellY(sheet, mergedRegion, moveNum);
                    isSkip[i] = true;
                }else if(cell != null && !StringUtils.isEmpty(cell.toString())){
                    moveCellY(sheet, rowIndex, cellIndex, moveNum);
                    isSkip[j] = true;
                }else if(cell == null){
                    row.createCell(cellIndex);
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
