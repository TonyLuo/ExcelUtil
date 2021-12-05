package com.github.tonyluo.excel;

import com.github.tonyluo.excel.util.ExcelConverter;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.List;

public class ExcelUtilTest {

    @Test
    public void exportToFile() {
//        Assert.assertEquals(ExcelConverter.columnName2Index("A"),1);

    }

    @Test
    public void exportToBytes() {
    }

    @Test
    public void importFromPath() throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException {
        List<Goods> list = ExcelUtil.importFromPath("src/test/resources/goods.xlsx", Goods.class, 1);
        for (Goods goods : list) {
            System.out.println(goods);

        }
        ExcelUtil.exportToFile("src/test/resources/export-goods.xlsx", list);

    }

    @Test
    public void testImportWithNotice() throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException {
        List<Goods> list = ExcelUtil.importFromPath("src/test/resources/goods-with-notice.xlsx", Goods.class, 2);
        for (Goods goods : list) {
            System.out.println(goods);

        }
        ExcelUtil.exportToFile("src/test/resources/export-goods-with-notice.xlsx", list);

    }

    @Test
    public void replaceBlankChar() {
        Assert.assertEquals("123", " 1 2     3".replaceAll("\\s*", ""));

    }

    @Test
    public void importFromInputStream() {
    }
}
