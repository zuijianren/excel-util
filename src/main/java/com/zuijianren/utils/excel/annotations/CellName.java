package com.zuijianren.utils.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用于声明标题名
 * <p>如果不声明, 则默认不处理</p>
 *
 * @author: 姜辞旧
 * @since: 2022/3/16 14:03
 * @version: 1.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellName {

    String value();
}
