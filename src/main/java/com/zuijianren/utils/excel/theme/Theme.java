package com.zuijianren.utils.excel.theme;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author: 姜辞旧
 * @since: 2022/3/16 15:25
 * @version: 1.0
 */
public abstract class Theme {

    protected Workbook workbook;

    /**
     * 顶部样式
     */
    protected CellStyle topStyle;

    /**
     * 标题样式
     */
    protected CellStyle titleStyle;

    /**
     * 内容样式
     */
    protected CellStyle contentStyle;


    /**
     * 由于 CellStyle 的创建必须要 Workbook, 因此需要一个初始化方法来设置 Workbook
     *
     * @param workbook Workbook对象
     */
    public void init(Workbook workbook) {
        // 初始化前, 清空原属性
        // 因为style对应单独的workbook
        // 至于为什么设计这个类, 是因为这样可以减少相关对象的创建
        this.topStyle = null;
        this.titleStyle = null;
        this.contentStyle = null;
        this.workbook = workbook;
    }


    /**
     * 获取顶部样式
     *
     * @return 标题样式
     */
    public abstract CellStyle getTopStyle();

    /**
     * 获取标题样式
     *
     * @return 标题样式
     */
    public abstract CellStyle getTitleStyle();

    /**
     * 获取内容样式
     *
     * @return 内容样式
     */
    public abstract CellStyle getContentStyle();

    /**
     * 是否创建最上方的顶行
     *
     * @return 是否需要
     */
    public boolean isCreateTop() {
        return true;
    }

    /**
     * 是否展示序号
     *
     * @return 是否展示序号
     */
    public boolean isShowNum() {
        return true;
    }
}
