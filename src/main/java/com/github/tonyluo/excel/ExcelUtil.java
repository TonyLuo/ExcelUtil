package com.github.tonyluo.excel;

import com.github.tonyluo.excel.annotation.ExcelSheet;
import com.github.tonyluo.excel.util.ExcelConverter;
import com.github.tonyluo.excel.util.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

public class ExcelUtil {


    /**
     * <p>导出Excel文件到磁盘</p>
     *
     * @param filePath   file path
     * @param entityList data list
     * @param <T>        class
     * @return file path
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> String exportToFile(String filePath, List<T> entityList) throws IllegalAccessException, InstantiationException, IOException, ClassNotFoundException {

        // workbook
        XSSFWorkbook workbook = ExcelConverter.generateWorkbook(entityList);

        FileOutputStream fileOutputStream = null;
        try {
            // workbook 2 FileOutputStream
            fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);

            fileOutputStream.flush();
        } finally {
            closeWorkbook(workbook, fileOutputStream);

        }

        return filePath;
    }


    /**
     * 导出Excel字节数据
     *
     * @param entityList data list
     * @param <T>        class
     * @return byte
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> byte[] exportToBytes(List<T> entityList) throws IllegalAccessException, InstantiationException, IOException, ClassNotFoundException {
        // workbook
        XSSFWorkbook workbook = ExcelConverter.generateWorkbook(entityList);

        ByteArrayOutputStream outputStream = null;
        byte[] result = null;

        try {
            outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);

            // flush
            outputStream.flush();
            result = outputStream.toByteArray();
        } finally {
            closeWorkbook(workbook, outputStream);
        }
        return result;

    }

    private static void closeWorkbook(XSSFWorkbook workbook, OutputStream outputStream) throws IOException {
        if (outputStream != null) {
            outputStream.flush();
            outputStream.close();
        }
        if (workbook != null) {
            // dispose of temporary files backing this workbook on disk
            workbook.close();
        }

    }


    /**
     * @param path     excel file path
     * @param clazz    class
     * @param startRow start row
     * @param <T>      class
     * @return java bean list
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importFromPath(String path, Class<T> clazz, int startRow) throws IOException, IllegalAccessException, InstantiationException, ClassNotFoundException {
        Workbook book = ExcelConverter.readFile(path);
        return importExcel(book, clazz, 0, startRow, -1);
    }

    /**
     * @param path  excel file path
     * @param clazz class
     * @param <T>   class
     * @return java bean list
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importFromPath(String path, Class<T> clazz) throws IOException, IllegalAccessException, InstantiationException, ClassNotFoundException {
        int startRow = getDefaultStartRow(clazz);
        return importFromPath(path, clazz, startRow);
    }

    /**
     * @param file     excel file path
     * @param clazz    class
     * @param startRow start row
     * @param <T>      class
     * @return java bean list
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importFromFile(File file, Class<T> clazz, int startRow) throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException {
        Workbook book = ExcelConverter.readFile(file);
        return importExcel(book, clazz, 0, startRow, -1);

    }

    /**
     * @param file  excel file path
     * @param clazz class
     * @param <T>   class
     * @return java bean list
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importFromFile(File file, Class<T> clazz) throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException {
        int startRow = getDefaultStartRow(clazz);
        return importFromFile(file, clazz, startRow);

    }

    private static <T> int getDefaultStartRow(Class<T> clazz) {
        int startRow = 1;
        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        if (excelSheet != null && StringUtils.isNotEmpty(excelSheet.notice())) {
            startRow = 2;

        }
        return startRow;
    }

    /**
     * @param stream   input stream
     * @param clazz    class
     * @param startRow start row
     * @param <T>      class
     * @return java bean list
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importFromInputStream(InputStream stream, Class<T> clazz, int startRow) throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException {
        Workbook book = ExcelConverter.readFromInputStream(stream);
        return importExcel(book, clazz, 0, startRow, -1);
    }

    /**
     * @param stream input stream
     * @param clazz  class
     * @param <T>    class
     * @return java bean list
     * @throws IOException            IOException
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importFromInputStream(InputStream stream, Class<T> clazz) throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException {
        int startRow = getDefaultStartRow(clazz);
        return importFromInputStream(stream, clazz, startRow);
    }


    //=====================================================================


    /**
     * @param book       workbook
     * @param clazz      class
     * @param sheetIndex sheet Index
     * @param startRow   start row
     * @param endRow     end row
     * @param <T>        class
     * @return java bean list
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     */
    public static <T> List<T> importExcel(Workbook book, Class<T> clazz, int sheetIndex, int startRow, int endRow) throws InstantiationException, IllegalAccessException, ClassNotFoundException {
        return ExcelConverter.getBeanListFromWorkBook(book, clazz, sheetIndex, startRow, endRow);
    }


}
