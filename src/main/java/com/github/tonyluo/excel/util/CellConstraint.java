package com.github.tonyluo.excel.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Map;

import static org.apache.poi.ss.usermodel.CellType.STRING;

public abstract class CellConstraint {
    public abstract DataValidationConstraint createConstraint(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol);

    public abstract Map getExplicitListValues();

    public boolean setAndFormatCellValue(Cell cell, Object valueObject) {
        String value = String.valueOf(this.getKeyByValue(valueObject));
        cell.setCellType(STRING);
        cell.setCellValue(value);
        return true;

    }

    public Object getAndFormatCellValue(Object valueObject) {
        return this.getValueByKey(valueObject);
    }

    private Object getKeyByValue(Object value) {
        Map hm = getExplicitListValues();
        for (Object o : hm.keySet()) {
            if (String.valueOf(hm.get(o)).equals(value)) {
                return o;
            }
        }
        return null;
    }

    private Object getValueByKey(Object key) {
        Map hm = getExplicitListValues();
        return hm.get(key);
    }

    public static CellConstraint getInstance(String constraintClass) throws ClassNotFoundException, IllegalAccessException, InstantiationException {

        return (CellConstraint) getClass(constraintClass).newInstance();
    }

    private static Class getClass(String constraintClass) throws ClassNotFoundException {

        return Class.forName(constraintClass);

    }
    public void addValidationData(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, XSSFDataValidationHelper dvHelper, DataValidationConstraint dvConstraint) {
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        XSSFDataValidation validation = (XSSFDataValidation)dvHelper.createValidation(
                dvConstraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

}
