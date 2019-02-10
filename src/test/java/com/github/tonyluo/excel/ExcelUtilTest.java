package com.github.tonyluo.excel;

import org.junit.Test;

import java.io.IOException;
import java.util.List;

public class ExcelUtilTest {

    @Test
    public void exportToFile() {
    }

    @Test
    public void exportToBytes() {
    }

    @Test
    public void importFromPath() throws IOException, InstantiationException, IllegalAccessException {
        List<Goods> list = ExcelUtil.importFromPath("src/test/resources/goods.xlsx", Goods.class,1);
        for (Goods goods : list) {
            System.out.println(goods);

        }
        ExcelUtil.exportToFile("src/test/resources/export-goods.xlsx", list);

    }

    @Test
    public void importFromFile() {
    }

    @Test
    public void importFromInputStream() {
    }
}
