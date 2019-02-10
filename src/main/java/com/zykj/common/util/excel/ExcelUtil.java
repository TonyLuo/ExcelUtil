package com.zykj.common.util.excel;

import com.zykj.common.util.excel.util.ExcelConverter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.List;

public class ExcelUtil {

    /**
     * 导出Excel文件到磁盘
     *
     * @param filePath
     * @param entityList 数据
     */
    public static <T> String exportToFile(String filePath, List<T> entityList) throws IllegalAccessException, InstantiationException, IOException {

        // workbook
        SXSSFWorkbook workbook = ExcelConverter.generateWorkbook(entityList);

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
     * @param entityList 数据
     */
    public static <T> byte[] exportToBytes(List<T> entityList) throws IllegalAccessException, InstantiationException, IOException {
        // workbook
        SXSSFWorkbook workbook = ExcelConverter.generateWorkbook(entityList);

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

    private static void closeWorkbook(SXSSFWorkbook workbook, OutputStream outputStream) throws IOException {
        if (outputStream != null) {
            outputStream.flush();
            outputStream.close();
        }
        if (workbook != null) {
            // dispose of temporary files backing this workbook on disk
            workbook.dispose();
        }

    }


    /**
     * generate javabean from excel path
     *
     * @param path
     * @param clazz
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> importFromPath(String path, Class<T> clazz, int startRow) throws IOException, IllegalAccessException, InstantiationException {
        Workbook book = ExcelConverter.readFile(path);
        return importExcel(book,clazz,0,startRow,-1);
    }

    /**
     * generate javabean from excel file
     *
     * @param file
     * @param clazz
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> importFromFile(File file, Class<T> clazz, int startRow) throws IOException, InstantiationException, IllegalAccessException {
        Workbook book = ExcelConverter.readFile(file);
        return importExcel(book,clazz,0,startRow,-1);

    }

    /**
     * generate javabean from InputStream
     *
     * @param stream
     * @param clazz
     * @param <T>
     * @return
     * @throws IOException
     * @throws InstantiationException
     * @throws IllegalAccessException
     */
    public static <T> List<T> importFromInputStream(InputStream stream, Class<T> clazz, int startRow) throws IOException, InstantiationException, IllegalAccessException {
        Workbook book = ExcelConverter.readFromInputStream(stream);
        return importExcel(book,clazz,0,startRow,-1);
    }


    //=====================================================================


    /**
     *
     * @param book
     * @param clazz
     * @param sheetIndex
     * @param startRow
     * @param endRow
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcel(Workbook book , Class<T> clazz,  int sheetIndex, int startRow, int endRow) throws InstantiationException, IllegalAccessException {
        return ExcelConverter.getBeanListFromWorkBook(book, clazz,sheetIndex, startRow,endRow);
    }


}
