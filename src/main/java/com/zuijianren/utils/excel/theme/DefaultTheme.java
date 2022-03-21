package com.zuijianren.utils.excel.theme;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author: 姜辞旧
 * @since: 2022/3/16 14:40
 * @version: 1.0
 */
public class DefaultTheme extends Theme {

    /**
     * 获取顶部样式
     *
     * @return 标题样式
     */
    @Override
    public CellStyle getTopStyle() {
        if (this.topStyle != null) {
            return this.topStyle;
        }
        CellStyle cellStyle = workbook.createCellStyle();
        // 水平垂直居中
        setCenter(cellStyle);
        // 设置字体
        Font titleFont = workbook.createFont();
        // 标题字体设置为粗体
        titleFont.setBold(true);
        cellStyle.setFont(titleFont);
        // 存入配置好的样式
        this.topStyle = cellStyle;
        return topStyle;
    }

    @Override
    public CellStyle getTitleStyle() {
        if (this.titleStyle != null) {
            return this.titleStyle;
        }
        CellStyle cellStyle = workbook.createCellStyle();
        // 水平垂直居中
        setCenter(cellStyle);
        // 填充样式
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 填充颜色
        cellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        // 设置字体
        Font titleFont = workbook.createFont();
        // 标题字体设置为粗体
        titleFont.setBold(true);
        cellStyle.setFont(titleFont);
        // 存入配置好的样式
        this.titleStyle = cellStyle;
        return titleStyle;
    }

    @Override
    public CellStyle getContentStyle() {
        if (this.contentStyle != null) {
            return this.contentStyle;
        }
        CellStyle cellStyle = workbook.createCellStyle();
        setCenter(cellStyle);
        this.contentStyle = cellStyle;
        return contentStyle;
    }


    /**
     * 设置居中
     *
     * @param cellStyle 单元格样式
     */
    private void setCenter(CellStyle cellStyle) {
        // 水平垂直居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }


}
