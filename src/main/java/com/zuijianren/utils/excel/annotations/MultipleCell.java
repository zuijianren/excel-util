package com.zuijianren.utils.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author: 姜辞旧
 * @since: 2022/3/17 11:30
 * @version: 1.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface MultipleCell {

    /**
     * 是否为基础类型
     */
    boolean basicType();

    /**
     * 如果为基础属性, 则需要设置该项, 否则不需要设置
     */
    String name() default "more";

    /**
     * 如果不是基础类型, 那么实际类型为
     */
    Class clazz();
}
