package com.zuijianren.utils.excel.theme;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author: 姜辞旧
 * @since: 2022/3/16 15:25
 * @version: 1.0
 */
public interface Theme {


    /**
     * 由于 CellStyle 的创建必须要 Workbook, 因此需要一个初始化方法来设置 Workbook
     *
     * @param workbook Workbook对象
     */
    void init(Workbook workbook);


    /**
     * 获取顶部样式
     *
     * @return 标题样式
     */
    CellStyle getTopStyle();

    /**
     * 获取标题样式
     *
     * @return 标题样式
     */
    CellStyle getTitleStyle();

    /**
     * 获取内容样式
     *
     * @return 内容样式
     */
    CellStyle getContentStyle();

    /**
     * 是否创建最上方的顶行
     *
     * @return 是否需要
     */
    boolean isCreateTop();

    /**
     * 是否展示序号
     *
     * @return 是否展示序号
     */
    boolean isShowNum();
}
