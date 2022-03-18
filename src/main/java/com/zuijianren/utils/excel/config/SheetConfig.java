package com.zuijianren.utils.excel.config;

import java.util.ArrayList;
import java.util.List;
import lombok.Data;

/**
 * sheet配置类
 *
 * @author 姜辞旧
 */
@Data
public class SheetConfig {

    /**
     * sheet名
     */
    private String sheetName;

    /**
     * 当存在一对多属性时, 存放该属性对应的列
     */
    private CellConfig multipleCell;

    /**
     * 各列属性
     */
    private final List<CellConfig> cellConfigList = new ArrayList<>();


}
