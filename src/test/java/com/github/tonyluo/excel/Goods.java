package com.github.tonyluo.excel;

import com.github.tonyluo.excel.annotation.ExcelCell;
import com.github.tonyluo.excel.annotation.ExcelSheet;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.time.Instant;
import java.util.Date;

@ExcelSheet(name="商品列表")
public class Goods {
    @ExcelCell(col="A",name="商品名",comment = "测试A1单元格备注功能") //
    private String name; //商品名

    @ExcelCell(col ="B",name="单位",width = 4, align = HorizontalAlignment.RIGHT,comment = "测试B1单元格备注功能")
    private String unit; //单位

    @ExcelCell(col ="C",name="规格",align = HorizontalAlignment.CENTER)
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
