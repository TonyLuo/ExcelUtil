package com.github.tonyluo.excel.util;

/**
 * @ClassName: ExcelConverter
 * @Description:
 * @author Tony
 * @date 2018年12月5日 下午1:19:56
 */

import com.github.tonyluo.excel.annotation.ExcelCell;
import com.github.tonyluo.excel.annotation.ExcelSheet;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * Usage: check Testcase ExcelConverterTest.java for detail
 */
public class ExcelConverter {


    public static Workbook readFile(File file) throws IOException {
        return readFromFile(file);

    }


    public static XSSFWorkbook readXSSFWorkbookFromInputStream(InputStream stream) throws IOException {
        return new XSSFWorkbook(stream);

    }

    public static Workbook readFromInputStream(InputStream stream) throws IOException {
        return WorkbookFactory.create(stream);

    }

    private static Workbook readFromFile(File file) throws IOException {
        return WorkbookFactory.create(file);

    }


    public static Workbook readFile(String path) throws IOException {
        File file = new File(path);
        if (!file.exists())
            throw new IOException("文件不存在");
        if (!file.isFile())
            throw new IOException("不是合法的文件");
        return readFile(file);
    }

    /**
     * @param name 'A', 'B',...,'AA','AB'
     * @return 0, 1, ..., 26, 27
     */
    public static int columnName2Index(String name) {
        int number = 0;
        if (StringUtils.isBlank(name)) {
            return number;
        }
        name = name.toUpperCase();
        for (int i = 0; i < name.length(); i++) {
            number = number * 26 + (name.charAt(i) - ('A' - 1));
        }
        return number - 1;
    }

    /**
     * @param index 0,1,...,26,27
     * @return 'A', 'B',...,'AA','AB'
     */
    public static String columnIndex2Name(int index) {
        StringBuilder sb = new StringBuilder();
        index++;
        while (index-- > 0) {
            sb.append((char) ('A' + (index % 26)));
            index /= 26;
        }
        return sb.reverse().toString();
    }

    private static <T extends Object> T convertBeanByRow(Row row, Class<T> clazz) throws IllegalAccessException, InstantiationException, ClassNotFoundException {


        T entity = clazz.newInstance();
        Field[] fields = clazz.getDeclaredFields();
        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);

        DataFormatter formatter = new DataFormatter();
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCell.class))
                continue;

            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            Integer col = getColumnIndex(excelSheet, field, excelCell);

            Cell cell = row.getCell(col);
            if (col == -1 || cell == null){
                continue;
            }


            // http://poi.apache.org/components/spreadsheet/quick-guide.html#CellContents
            Object cellValue = null;
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getRichStringCellValue().getString();

                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = cell.getDateCellValue();

                    } else {
                        //https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Cell.html#setCellType%28int%29
                        //If what you want to do is get a String value for your numeric cell, stop!. This is not the way to do it. Instead, for fetching the string value of a numeric or boolean or date cell, use DataFormatter instead.
                        cellValue = formatter.formatCellValue(cell);
//                        cellValue = cell.getNumericCellValue();
                    }
                    break;
                case BOOLEAN:
                    cellValue = cell.getBooleanCellValue();

                    break;
                case FORMULA:
                    cellValue = cell.getRichStringCellValue().getString();

                    break;
                case BLANK:
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellValue = cell.getRichStringCellValue().getString();
            }


            field.setAccessible(true);
            String constraintClass = excelCell.constraintClass();
            if (StringUtils.isNotEmpty(constraintClass)) {
                CellConstraint cellConstraint = CellConstraint.getInstance(constraintClass);
                cellValue = cellConstraint.getAndFormatCellValue(cellValue);
            }

            field.set(entity, FieldParser.parseValue(field, cellValue));


        }
        return entity;
    }

    private static Integer getColumnIndex(ExcelSheet excelSheet, Field field, ExcelCell excelCell) {
        Integer col = getColumnIndex(excelSheet, field);
        if (null == col) {
            col = getColumnIndex(excelCell);
        }
        return col;
    }

    /**
     *
     * @param excelSheet
     * @param field
     * @return  如果ExcelSheet.cols 没有配置返回null；如果ExcelSheet.cols有值，但是字段不在cols里面，返回-1
     */
    private static Integer getColumnIndex(ExcelSheet excelSheet, Field field) {
        if (null == excelSheet || null == field) {
            return null;
        }
        String[] cols = excelSheet.cols();
        if (null == cols || cols.length < 1) {
            return null;
        }

        for (int i = 0; i < cols.length; i++) {
            if (field.getName().equals(cols[i])) {
                return i;
            }

        }
        return -1;
    }

    private static int getColumnIndex(ExcelCell excelCell) {
//        int col = excelCell.col();
//        int columnIndex = 0;
//        String colName = excelCell.col();
//        if (StringUtils.isNotEmpty(colName)) {
//            columnIndex = columnName2Index(colName);
//        }
//        return columnIndex;
        return columnName2Index(excelCell.col());
    }


    public static <T> List<T> getBeanListFromWorkBook(Workbook book, Class<T> clazz, int startRow) throws IllegalAccessException, InstantiationException, ClassNotFoundException {
        //start from second row, skip the header
        return getBeanListFromWorkBook(book, clazz, startRow, -1);


    }

    /**
     * @param book     Workbook
     * @param clazz    class
     * @param startRow <p>start row</p>
     * @param endRow   <p>end row, if endRow = -1, will get the last row of the sheet</p>
     * @param <T>      class
     * @return java bean list
     * @throws IllegalAccessException IllegalAccessException
     * @throws InstantiationException InstantiationException
     */
    public static <T> List<T> getBeanListFromWorkBook(Workbook book, Class<T> clazz, int startRow, int endRow) throws IllegalAccessException, InstantiationException, ClassNotFoundException {

        return getBeanListFromWorkBook(book, clazz, 0, startRow, endRow);
    }

    /**
     * @param book       Workbook
     * @param clazz      class
     * @param sheetIndex sheetIndex
     * @param startRow   <p>start row</p>
     * @param endRow     <p>end row, if endRow = -1, will get the last row of the sheet</p>
     * @param <T>        class
     * @return java bean list
     * @throws IllegalAccessException IllegalAccessException
     * @throws InstantiationException InstantiationException
     */
    public static <T> List<T> getBeanListFromWorkBook(Workbook book, Class<T> clazz, int sheetIndex, int startRow, int endRow) throws IllegalAccessException, InstantiationException, ClassNotFoundException {
        List<T> list = new ArrayList<T>();
        Sheet sheet = book.getSheetAt(sheetIndex);
        if (endRow < 0) {
            endRow = sheet.getLastRowNum();
        }
        for (int i = startRow; i <= endRow; i++) {
            T t = convertBeanByRow(sheet.getRow(i), clazz);
            list.add(t);
        }
        return list;
    }

    private static void setCellStyleAndValue(Object valueObject, Field field, Workbook workbook, Cell cell) throws IllegalAccessException, InstantiationException, ClassNotFoundException {

        ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
        if (null == excelCell) {
            return;
        }

        Class<?> fieldType = field.getType();
        CellStyle cellStyle = setCellStyle(workbook, fieldType, cell, excelCell);

        if (valueObject == null) {
            cell.setCellValue("");
            return;
        }


        String value = StringUtils.trimToEmpty(valueObject.toString());
        String constraintClass = excelCell.constraintClass();
        if (StringUtils.isNotEmpty(constraintClass)) {

            CellConstraint cellConstraint = CellConstraint.getInstance(constraintClass);
            if (cellConstraint.setAndFormatCellValue(cell, valueObject)) {
                return;
            }

        }
        cell.setCellType(NUMERIC);

        if (String.class.equals(fieldType)) {
            cell.setCellType(STRING);
            cell.setCellValue(value);
        } else if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            cell.setCellType(BOOLEAN);
            cell.setCellValue(FieldParser.parseBoolean(value));

        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            Short shortValue = FieldParser.parseShort(value);
            if (null != shortValue) {
                cell.setCellValue(shortValue);

            }

        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            Integer intValue = FieldParser.parseInt(value);
            if (null != intValue) {
                cell.setCellValue(intValue);

            }

        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            Long longValue = FieldParser.parseLong(value);
            if (null != longValue) {
                cell.setCellValue(longValue);

            }

        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            Float floatValue = FieldParser.parseFloat(value);
            if (null != floatValue) {
                cell.setCellValue(floatValue);

            }

        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            Double doubleValue = FieldParser.parseDouble(value);
            if (null != doubleValue) {
                cell.setCellValue(doubleValue);

            }

        } else if (Date.class.equals(fieldType) || Instant.class.equals(fieldType)) {
            Date date = FieldParser.parseDate(value, excelCell);
            if (null != date) {
//                CreationHelper createHelper = workbook.getCreationHelper();
//                cellStyle.setDataFormat(
//                    createHelper.createDataFormat().getFormat("yyyy/m/d"));
                cell.setCellValue(date);
            }

        } else {
            cell.setCellType(STRING);
            cell.setCellValue(value);
        }

//        cell.setCellStyle(cellStyle);


    }

    private static CellStyle setCellStyle(Workbook workbook, Class<?> fieldType, Cell cell, ExcelCell excelCell) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(excelCell.wrapText()); //Set wordwrap
        cellStyle.setLocked(excelCell.locked()); //Set locked
        if (excelCell.locked()) {
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cellStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        cellStyle.setAlignment(excelCell.align());

        if (Date.class.equals(fieldType) || Instant.class.equals(fieldType)) {
            cell.setCellType(NUMERIC);

            //http://poi.apache.org/components/spreadsheet/quick-guide.html#CreateDateCells
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(excelCell.dateFormat()));

        }
        if (StringUtils.isNotEmpty(excelCell.format())) {
            DataFormat format = workbook.createDataFormat();
            cellStyle.setDataFormat(format.getFormat(excelCell.format()));

        }


        cell.setCellStyle(cellStyle);

        return cellStyle;
    }

    protected static <T> void setRowWithBean(Workbook workbook, XSSFSheet sheet, Row row, T entity, boolean isHeader) throws IllegalAccessException, InstantiationException, ClassNotFoundException {
        Field[] fields = entity.getClass().getDeclaredFields();
        ExcelSheet excelSheet = entity.getClass().getAnnotation(ExcelSheet.class);

        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCell.class)) {
                continue;
            }
            field.setAccessible(true);
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            Integer col = getColumnIndex(excelSheet, field, excelCell);
            if (col == -1 ){
                continue;
            }
            if (row != null) {
                Cell cell = row.createCell(col);

                if (isHeader) {
                    setSheetHeader(workbook, sheet, entity, field, excelCell, cell);

                } else {

                    Object fieldValue = field.get(entity);
                    String valueObject = FieldParser.formatValue(field, fieldValue);


                    setCellStyleAndValue(valueObject, field, workbook, cell);

                }
            }


        }

    }

    private static <T> void setSheetHeader(Workbook workbook, XSSFSheet sheet, T entity, Field field, ExcelCell excelCell, Cell cell) {
        String colName = excelCell.name();
        if (StringUtils.isEmpty(colName)) {
            colName = field.getName();
        }

        //增加备注
        String commentText = excelCell.comment();
        if (StringUtils.isNotEmpty(commentText)) {
            CreationHelper createHelper = workbook.getCreationHelper();

            Drawing drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = createHelper.createClientAnchor();

            //fix bug: https://bz.apache.org/bugzilla/show_bug.cgi?id=59393
            anchor.setRow1(cell.getRowIndex());
            anchor.setCol1(cell.getColumnIndex());
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(createHelper.createRichTextString(commentText));

            // Assign the comment to the cell
            cell.setCellComment(comment);
        }
        // head-style、field-data-style
        CellStyle headerCellStyle = workbook.createCellStyle();
        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        if (excelCell.required()) {
            headerFont.setColor(IndexedColors.RED.getIndex());
        }
        //headerFont.setColor(IndexedColors.RED.getIndex());
        // Create a CellStyle with the font
        headerCellStyle.setFont(headerFont);
        int headColorIndex = -1;
        ExcelSheet excelSheet = entity.getClass().getAnnotation(ExcelSheet.class);
        if (excelSheet != null) {
            headColorIndex = excelSheet.headColor().getIndex();
        }
        if (headColorIndex > -1) {
            headerCellStyle.setFillForegroundColor((short) headColorIndex);
            headerCellStyle.setFillBackgroundColor((short) headColorIndex);
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        headerCellStyle.setWrapText(excelCell.wrapText()); //Set wordwrap
        cell.setCellStyle(headerCellStyle);
        cell.setCellValue(colName);
    }

    private static <T> void createSheetNotice(Workbook workbook, Sheet sheet, Row row, ExcelSheet excelSheet, T entity) {
        Cell cell = row.createCell(0);

        CellStyle noticeCellStyle = workbook.createCellStyle();
        // Create a Font for styling header cells
        Font noticeFont = workbook.createFont();
        noticeFont.setBold(true);
        noticeFont.setFontHeightInPoints((short) 16);

        // Create a CellStyle with the font
        noticeCellStyle.setFont(noticeFont);
        noticeCellStyle.setWrapText(true); //Set wordwrap
        noticeCellStyle.setAlignment(HorizontalAlignment.LEFT);
        noticeCellStyle.setVerticalAlignment(VerticalAlignment.TOP);

        cell.setCellStyle(noticeCellStyle);

        //set value
        String notice = excelSheet.notice();
        cell.setCellValue(notice);

        //merge cell for the notice
        Field[] fields = entity.getClass().getDeclaredFields();
        int lastCol = 0;
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCell.class)) {
                continue;
            }
            lastCol++;
        }

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCol - 1));

        //set row height
        String regex = "\\r?\\n";
        String noticeList[] = notice.split(regex);
        int lineCount = noticeList.length + 1;
        if (lineCount > 1) {
            short fontHeight = noticeFont.getFontHeightInPoints();
//            lineCount = lineCount + 3;
            short finalRowHeight = (short) (fontHeight * lineCount * 1.4);
            row.setHeightInPoints(finalRowHeight);

        }
    }


    protected static <T> Object[] getRowByBean(T entity) throws IllegalAccessException {
        Field[] fields = entity.getClass().getDeclaredFields();
        ExcelSheet excelSheet = entity.getClass().getAnnotation(ExcelSheet.class);
        int columnLength = 0;
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelCell.class)) {
                columnLength++;
            }
        }
        Object[] list = new Object[columnLength];


        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCell.class)){
                continue;
            }


            field.setAccessible(true);
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);

            Integer col = getColumnIndex(excelSheet, field, excelCell);
            if (col == -1 ){
                continue;
            }
            Object fieldValue = field.get(entity);
            if (fieldValue == null) {
                continue;
            }

            list[col] = FieldParser.formatValue(field, fieldValue);


        }
        return list;

    }

    protected static <T> String[] getRowHeaderByClass(Class<T> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        int columnLength = 0;
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelCell.class)) {
                columnLength++;
            }
        }
        String[] list = new String[columnLength];

        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCell.class))
                continue;

            field.setAccessible(true);
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            Integer col = getColumnIndex(excelSheet, field, excelCell);
            if (col == -1 ){
                continue;
            }
            String colName = excelCell.name();
            if (StringUtils.isEmpty(colName)) {
                colName = field.getName();
            }

            list[col] = colName;


        }
        return list;

    }

    public static <T> XSSFWorkbook generateWorkbook(List<T> entityList) throws InstantiationException, IllegalAccessException, ClassNotFoundException {

        if (entityList == null || entityList.size() == 0) {
            return null;
        }

        // Create a Workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
//        XSSFWorkbook workbook = new XSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
//        workbook.setCompressTempFiles(true); // temp files will be gzipped

        T entity = entityList.get(0);
        XSSFSheet sheet = createSheet(workbook, entity);

        int startRow = createSheetHeader(workbook, sheet, entity);
//        sheet.trackAllColumnsForAutoSizing();
        // Create Other rows and cells with employees data

        int rowNumber = startRow;
        for (T e : entityList) {
            Row row = sheet.createRow(rowNumber++);
            setRowWithBean(workbook, sheet, row, e, false);

        }

        // reset sheet
        reArrangeSheet(sheet, entity, startRow);

        return workbook;
    }

    private static <T> void reArrangeSheet(XSSFSheet sheet, T entity, int firstRow) throws IllegalAccessException, InstantiationException, ClassNotFoundException {
        Field[] fields = entity.getClass().getDeclaredFields();
        ExcelSheet excelSheet = entity.getClass().getAnnotation(ExcelSheet.class);

        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCell.class)) {
                continue;
            }
            field.setAccessible(true);
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            Integer col = getColumnIndex(excelSheet, field, excelCell);
            if (col == -1 ){
                continue;
            }
            int width = excelCell.width();
            if (width > -1) {
                if (width > 0) {
                    width = (width + 2) * 256;
                }
                // 1/256th of a character width
                sheet.setColumnWidth(col, width);

            } else {
                sheet.autoSizeColumn(col);
            }

            String constraintClass = excelCell.constraintClass();
            if (StringUtils.isNotEmpty(constraintClass)) {
                CellConstraint cellConstraint = CellConstraint.getInstance(constraintClass);
                cellConstraint.createConstraint(sheet, firstRow, sheet.getLastRowNum() + 100, col, col);
            }
            if (excelCell.hidden()) {
                sheet.setColumnHidden(col, true);
            }
        }
    }


    private static <T> int createSheetHeader(XSSFWorkbook workbook, XSSFSheet sheet, T entity) throws InstantiationException, IllegalAccessException, ClassNotFoundException {
        ExcelSheet excelSheet = entity.getClass().getAnnotation(ExcelSheet.class);
        int startRow = 0;
        if (excelSheet != null && StringUtils.isNotEmpty(StringUtils.trimToEmpty(excelSheet.notice()))) {
            Row noticeRow = sheet.createRow(startRow);
            startRow++;
            createSheetNotice(workbook, sheet, noticeRow, excelSheet, entity);
        }
        // Create a Row
        Row headerRow = sheet.createRow(startRow);
        startRow++;
        // Create cells
        setRowWithBean(workbook, sheet, headerRow, entity, true);


        return startRow;
    }

    private static <T> XSSFSheet createSheet(XSSFWorkbook workbook, T entity) {
        String sheetName = entity.getClass().getSimpleName();
        ExcelSheet excelSheet = entity.getClass().getAnnotation(ExcelSheet.class);
        if (excelSheet != null) {
            if (StringUtils.isNotBlank(excelSheet.name())) {
                sheetName = excelSheet.name().trim();
            }

        }

        XSSFSheet existSheet = workbook.getSheet(sheetName);
        if (existSheet != null) {
            for (int i = 2; i <= 1000; i++) {
                String newSheetName = sheetName.concat(String.valueOf(i));  // avoid sheetName duplicate
                existSheet = workbook.getSheet(newSheetName);
                if (existSheet == null) {
                    sheetName = newSheetName;
                    break;
                } else {
                    continue;
                }
            }
        }

        XSSFSheet sheet = workbook.createSheet(sheetName);

        if (excelSheet != null) {
            int colSplit = excelSheet.colSplit();
            int rowSplit = excelSheet.rowSplit();
            int leftmostColumn = excelSheet.leftmostColumn();
            int topRow = excelSheet.topRow();
            if (colSplit > -1 && rowSplit > -1 && leftmostColumn > -1 && topRow > -1) {
                sheet.createFreezePane(colSplit, rowSplit, leftmostColumn, topRow);

            } else if (colSplit > -1 && rowSplit > -1) {
                sheet.createFreezePane(colSplit, rowSplit);
            }

            if (excelSheet.protectSheet()) {
                sheet.protectSheet(excelSheet.protectSheetPassword());
            }


        }
        return sheet;
    }

}
