package com.zuijianren.utils.excel.config;

import java.lang.reflect.Field;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @author: 姜辞旧
 * @since: 2022/3/16 16:31
 * @version: 1.0
 */
@Data
@AllArgsConstructor
public class CellConfig {

    /**
     * 标题行名字
     */
    private String cellName;

    /**
     * 用于获取value的field
     */
    private Field field;

    /**
     * 类型
     */
    private Class<?> clazz;

    /**
     * 是否为 "多"
     */
    private boolean more;

    /**
     * 是否为内嵌属性
     */
    private boolean nestAttr;

    public CellConfig(Field field, Class<?> clazz) {
        this.field = field;
        this.clazz = clazz;
    }
}
