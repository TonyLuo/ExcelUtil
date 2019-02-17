package com.github.tonyluo.excel;

import com.github.tonyluo.excel.util.CellConstraint;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

public class GoodsConstraint {
    public static final class TypeConstraint extends CellConstraint {
        private static Map<String, Integer> values = new HashMap<String, Integer>(){{
            put("中药",1);
            put("西药",2);
        }};

        @Override
        public DataValidationConstraint createConstraint(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {

            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
            DataValidationConstraint dvConstraint =
//                    dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});
                    dvHelper.createExplicitListConstraint(Arrays.stream(values.keySet().toArray()).toArray(String[]::new));
            addValidationData(sheet, firstRow, lastRow, firstCol, lastCol, dvHelper, dvConstraint);

            return  dvConstraint;
        }

        @Override
        public Map getExplicitListValues() {
            return values;
        }
    }

    public static final class SellPriceConstraint extends CellConstraint {
        private static Map<String, Integer> values = new HashMap<String, Integer>(){{
//            put("min",10);
//            put("max",100);
        }};
        @Override
        public DataValidationConstraint createConstraint(XSSFSheet sheet,int firstRow,int lastRow,int firstCol,int lastCol) {

            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
            XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                    dvHelper.createNumericConstraint(
                            XSSFDataValidationConstraint.ValidationType.DECIMAL,
                            XSSFDataValidationConstraint.OperatorType.BETWEEN,
                            "10", "100");
            addValidationData(sheet, firstRow, lastRow, firstCol, lastCol, dvHelper, dvConstraint);
            return  dvConstraint;

        }


        @Override
        public Map getExplicitListValues() {
            return values;
        }

        @Override
        public Object getAndFormatCellValue(Object valueObject) {
            return valueObject;
        }
        @Override
        public boolean setAndFormatCellValue(Cell cell, Object valueObject) {
            return false;

        }
    }


}
