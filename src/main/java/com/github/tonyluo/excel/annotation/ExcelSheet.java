package com.github.tonyluo.excel.annotation;


import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.*;

/**
 * 表信息
 *
 * @author Tony 2018-09-08 20:51:26
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {

    /**
     * 表名称
     *
     * @return String
     */
    String name() default "";

    /**
     * 表头/首行的颜色
     *
     * @return HSSFColorPredefined
     */
    HSSFColor.HSSFColorPredefined headColor() default HSSFColor.HSSFColorPredefined.LIGHT_GREEN;

    /**
     * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
     * <p>
     *     If both colSplit and rowSplit are zero then the existing freeze pane is removed
     * </p>
     * @param colSplit      Horizontal position of split.
     * @param rowSplit      Vertical position of split.
     * @param leftmostColumn   Left column visible in right pane.
     * @param topRow        Top row visible in bottom pane
     */

    /**
     * colSplit Horizontal position of split.
     *
     * @return int
     */
    int colSplit() default -1;

    int rowSplit() default -1;

    int leftmostColumn() default -1;

    int topRow() default -1;

    //注：要锁定单元格需先为此表单设置保护密码，设置之后此表单默认为所有单元格锁定，可使用setLocked(false)为指定单元格设置不锁定。
    boolean protectSheet() default false;

    String protectSheetPassword() default "";

    /**
     * 字段名称排序，excel优先按照cols里面的字段排序，减少调整字段顺序操作痛苦
     *
     * @return
     */
    String[] cols() default {};

    String notice() default "填写须知：\n1、不能增加、删除列；\n2、不能修改灰色单元格；\n3、红色字段为必填字段，黑色字段为选填字段；\n";


}
