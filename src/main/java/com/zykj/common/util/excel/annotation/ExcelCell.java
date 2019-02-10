package com.zykj.common.util.excel.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.*;

/**
 * @ClassName: ExcelCell
 * @Description: 实体字段与excel列号关联的注解
 * @author Tony
 *
 */
@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCell {
//    int col() default 0; // excel value column
    String col()  default ""; // excel column 'A' ,'B' ... 'AA'

    String name() default ""; // excel  column name

    Class<?> Type() default String.class; // excel cell value type


    /**
     * 列宽 (大于0时生效; 如果不指定列宽，将会自适应调整宽度；)
     * Set the width (character width)<p>
     * @return int
     */
    int width() default -1;

    /**
     * 水平对齐方式
     *
     * @return HorizontalAlignment
     */
    HorizontalAlignment align() default HorizontalAlignment.LEFT;
    /**
     * 时间格式化，日期类型时生效
     *
     * @return String
     */
    String dateFormat() default "yyyy-MM-dd";


    /**
     * format: "#,##0.0000"
     * @return
     */
    String format()  default "";

    /**
     * 单元格备注
     *
     * @return String
     */
    String comment() default "";

    boolean wrapText() default false;
    boolean hidden() default false;


    /**
     *
     *     //注：要锁定单元格需先为此表单设置保护密码，设置之后此表单默认为所有单元格锁定，可使用setLocked(false)为指定单元格设置不锁定。
     * @return
     */
    boolean locked() default false; // ExcelSheet的protectSheet为true时候才能生效



//    String[] dvConstraint() default {};



}

