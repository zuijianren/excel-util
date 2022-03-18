package com.zuijianren.utils.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用于声明sheet的名字
 *
 * @author: 姜辞旧
 * @since: 2022/3/16 14:03
 * @version: 1.0
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface SheetName {

    String value() default "主表";
}
