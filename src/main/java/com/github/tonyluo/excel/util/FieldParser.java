package com.github.tonyluo.excel.util;


import com.github.tonyluo.excel.annotation.ExcelCell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;


public final class FieldParser {

    private static final String DEFAULT_DATE_PATTERN = "EEE MMM dd HH:mm:ss z yyyy";
    private static Set<String> datePatternSet = null;

    private FieldParser() {

    }

    private static Set<String> initDatePatternSet(String datePattern) {
        if (null == datePatternSet) {
            datePatternSet = new HashSet<>();
            datePatternSet.add(DEFAULT_DATE_PATTERN);
            datePatternSet.add("m/d/yyyy");
            datePatternSet.add("MM/dd/yyyy");
            datePatternSet.add("yyyyMMdd");
            datePatternSet.add("yyyy-MM-dd");
        }

        if (StringUtils.isNotEmpty(datePattern)) {
            datePatternSet.add(datePattern);

        }
        return datePatternSet;
    }

    public static Byte parseByte(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }
        try {
            return Byte.valueOf(value);
        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseByte] illegal input=" + value, e);
        }
    }

    public static Boolean parseBoolean(String value) {
        value = trimValue(value);
        if (StringUtils.isEmpty(value)) {
            return null;
        }
        if ("Y".equalsIgnoreCase(value) ||
                "YES".equalsIgnoreCase(value) ||
                "1".equalsIgnoreCase(value) ||
                Boolean.TRUE.toString().equalsIgnoreCase(value)) {
            return Boolean.TRUE;
        } else {
            return Boolean.FALSE;
        }
    }

    public static Integer parseInt(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }
        try {
            value = value.replaceAll("　", "");
            return new BigDecimal(value).setScale(0, RoundingMode.HALF_UP).intValueExact();
        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseInt] illegal input. value=" + value, e);
        }
    }

    public static Short parseShort(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }
        try {
            return new BigDecimal(value).setScale(0, RoundingMode.HALF_UP).shortValueExact();
        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseShort] illegal input. value=" + value, e);
        }
    }

    public static Long parseLong(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }
        try {
            return new BigDecimal(value).setScale(0, RoundingMode.HALF_UP).longValueExact();

        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseLong] illegal input. value=" + value, e);
        }
    }

    private static boolean trimAndVerifyValue(String value) {
        value = trimValue(value);
        return NumberUtils.isCreatable(value);
    }

    private static String trimValue(String value) {
        value = StringUtils.trimToEmpty(value);
        value = value.replaceAll("\\s*", "");
        return value;
    }

    public static Float parseFloat(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }
        try {
            return Float.valueOf(value);
        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseFloat] illegal input. value=" + value, e);
        }
    }

    public static Double parseDouble(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }

        try {
            return Double.valueOf(value);
        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseDouble]illegal input. value=" + value, e);
        }
    }

    public static BigDecimal parseBigDecimal(String value) {
        if (!trimAndVerifyValue(value)) {
            return null;
        }
        try {
            return new BigDecimal(value);
        } catch (NumberFormatException e) {
            throw new RuntimeException("[parseBigDecimal]illegal input. value=" + value, e);
        }
    }

    public static Instant parseInstant(String value, ExcelCell excelCell) {
        value = StringUtils.trimToEmpty(value);
        if (StringUtils.isEmpty(value)) {
            return null;
        }

        try {
            if (NumberUtils.isCreatable(value)) {
                value = trimValue(value);
                double doubleValue = Double.parseDouble(value);
                Date date = DateUtil.getJavaDate(doubleValue);
                return date.toInstant();
            } else {
                return parseInstantFromString(value, excelCell);
            }

        } catch (NumberFormatException e) {
            return parseInstantFromString(value, excelCell);

        }

    }

    public static Instant parseInstantFromString(String value, ExcelCell excelCell) {
        if (StringUtils.isBlank(value)) {
            return null;
        }
        DateTimeFormatterBuilder dateTimeFormatterBuilder = new DateTimeFormatterBuilder();
//            .appendOptional(DateTimeFormatter.ofPattern(DEFAULT_DATE_PATTERN));
//        datePatternSet.forEach(pattern ->{
//            dateTimeFormatterBuilder.appendOptional(DateTimeFormatter.ofPattern(pattern));
//        });
        String datePattern = null;
        if (excelCell != null) {
            //DateTimeFormatterBuilder appendOptional 不能重复添加相同的值
            datePattern = excelCell.dateFormat();
//            if (!DEFAULT_DATE_PATTERN.equals(datePattern)) {
//                dateTimeFormatterBuilder.appendOptional(DateTimeFormatter.ofPattern(datePattern));
//            }


        }
        initDatePatternSet(datePattern).forEach(pattern -> {
            dateTimeFormatterBuilder.appendOptional(DateTimeFormatter.ofPattern(pattern));
        });
//        DateUtils.parseDate(value,initDatePatternSet(datePattern).to())

        DateTimeFormatter dfs = dateTimeFormatterBuilder.toFormatter();
        LocalDate localDate = LocalDate.parse(value, dfs);
        ZoneId zoneId = ZoneId.systemDefault();
        ZonedDateTime zdt = localDate.atStartOfDay(zoneId);
        return zdt.toInstant();

    }

    public static Date parseDate(String value, ExcelCell excelCell) {
        if (StringUtils.isBlank(value)) {
            return null;
        }
        Instant instant = parseInstant(value, excelCell);
        return Date.from(instant);
    }


    /**
     * 参数解析 （支持：Byte、Boolean、String、Short、Integer、Long、Float、Double、Date）
     *
     * @param field       field
     * @param valueObject value
     * @return Object
     */
    public static Object parseValue(Field field, Object valueObject) {
        if (valueObject == null) {
            return null;
        }
        Class<?> fieldType = field.getType();

        ExcelCell excelCell = field.getAnnotation(ExcelCell.class);

        String value = valueObject.toString().trim();

        if (String.class.equals(fieldType)) {
            return value;
        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            return parseShort(value);
        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            return parseInt(value);
        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            return parseLong(value);
        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            return parseFloat(value);
        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            return parseDouble(value);
        } else if (BigDecimal.class.equals(fieldType)) {
            return parseBigDecimal(value);
        } else if (Date.class.equals(fieldType)) {
            return parseDate(value, excelCell);
        } else if (Instant.class.equals(fieldType)) {
            return parseInstant(value, excelCell);
        } else if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            return parseBoolean(value);
        } else if (Byte.class.equals(fieldType) || Byte.TYPE.equals(fieldType)) {
            return parseByte(value);
        } else {
            throw new RuntimeException("[parseValue] illegal type, type=" + fieldType);
        }
    }

    /**
     * 参数格式化为String
     *
     * @param field field
     * @param value value
     * @return String
     */
    public static String formatValue(Field field, Object value) {
        if (value == null || null == field) {
            return null;
        }

        Class<?> fieldType = field.getType();

        if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (String.class.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            return String.valueOf(value);
        } else if (Date.class.equals(fieldType)) {

            return String.valueOf(DateUtil.getExcelDate((Date) value));
        } else if (Instant.class.equals(fieldType)) {

            return String.valueOf(DateUtil.getExcelDate(Date.from((Instant) value)));
        } else {
            throw new RuntimeException("[formatValue] illegal type, type=" + fieldType);
        }
    }

}
