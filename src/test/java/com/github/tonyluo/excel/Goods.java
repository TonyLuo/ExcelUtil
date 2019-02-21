package com.github.tonyluo.excel;

import com.github.tonyluo.excel.annotation.ExcelCell;
import com.github.tonyluo.excel.annotation.ExcelSheet;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.time.Instant;
import java.util.Date;

@ExcelSheet(name="商品列表",notice = "填写须知：\n" +
        "1、不能增加、删除列；\n" +
        "2、不能修改灰色单元格；\n" +
        "3、红色字段为必填字段，黑色字段为选填字段；\n" +
        "4、删除餐厅编码后再导入，系统会根据ID删除没有餐厅编码的记录；\n" +
        "5、删除餐厅编码后再导入，系统会根据ID删除没有餐厅编码的记录；\n" +
        "6、删除餐厅编码后再导入，系统会根据ID删除没有餐厅编码的记录；删除餐厅编码后再导入，系统会根据ID删除没有餐厅编码的记录删除餐厅编码后再导入，系统会根据ID删除没有餐厅编码的记录\n" +
        "7、删除餐厅编码后再导入，系统会根据ID删除没有餐厅编码的记录；")
public class Goods {
    @ExcelCell(col="A",name="商品名",comment = "测试A1单元格备注功能",locked = true,required = true) //
    private String name; //商品名

    @ExcelCell(col ="B",name="单位",width = 4, align = HorizontalAlignment.RIGHT,comment = "测试B1单元格备注功能")
    private String unit; //单位

    @ExcelCell(col ="C",name="规格",required = true, align = HorizontalAlignment.CENTER)
    private String format; //规格

    @ExcelCell(col ="D",name="生产厂家", wrapText= true, width = 4, comment = "测试单元格宽度、自动换行、备注功能")
    private String factory;//生产厂家

    @ExcelCell(col ="E",name="生产时间", dateFormat = "yyyy-MM-dd HH:mm:ss")
    private Date manufactureTime;

    @ExcelCell(col="F", name="出厂日期",dateFormat = "MM/dd/yyyy")
    private Instant productionDate;

    @ExcelCell(col="G", name="数量", comment = "测试G1单元格备注功能")
    private int quantity;

    @ExcelCell(col="H", name="价格",hidden = true )
    private double price;

    @ExcelCell(col="I", name="售价",format ="#,##0.00",
            constraintClass="com.github.tonyluo.excel.GoodsConstraint$SellPriceConstraint")
    private Float sellPrice;

    @ExcelCell(col="J", name="类型",
            constraintClass="com.github.tonyluo.excel.GoodsConstraint$TypeConstraint")
    private Integer type;

    @Override
    public String toString() {
        return "Goods{" +
            "name='" + name + '\'' +
            ", unit='" + unit + '\'' +
            ", format='" + format + '\'' +
            ", factory='" + factory + '\'' +
            '}';
    }
}
