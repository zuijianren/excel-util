package com.zuijianren.utils.excel.core;

import com.zuijianren.utils.excel.annotations.CellName;
import com.zuijianren.utils.excel.annotations.MultipleCell;
import com.zuijianren.utils.excel.annotations.SheetName;
import com.zuijianren.utils.excel.config.CellConfig;
import com.zuijianren.utils.excel.config.SheetConfig;
import com.zuijianren.utils.excel.exceptions.AnnotationException;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
import lombok.extern.java.Log;

/**
 * @author: 姜辞旧
 * @since: 2022/3/17 14:25
 * @version: 1.0
 */
@Log
public class ConfigParse {

    /**
     * 配置
     */
    private final SheetConfig sheetConfig = new SheetConfig();


    /**
     * 解析注解
     * <p>初步解析注解, 用于更进一步的解析实体类, 生成配置</p>
     *
     * @param targetClass 目标类
     */
    public SheetConfig parseAnnotation(Class<?> targetClass) {
        if (!targetClass.isAnnotationPresent(SheetName.class)) {
            throw new AnnotationException(
                String.format(
                    "在 %s 类中, 没有找到 com.zuijianren.utils.excel.annotations.SheetName 注解, 导出失败",
                    targetClass.getName())
            );
        }

        log.info("解析" + targetClass.getName());

        String sheetName = targetClass.getAnnotation(SheetName.class).value();

        // 设置表名
        sheetConfig.setSheetName(sheetName);

        for (Field declaredField : targetClass.getDeclaredFields()) {
            // 跳过未声明CellName或者MultipleCell注解的属性
            if (!declaredField.isAnnotationPresent(CellName.class) &&
                !declaredField.isAnnotationPresent(MultipleCell.class)) {
                continue;
            }
            parseField(declaredField);
        }

        return sheetConfig;
    }

    /**
     * 解析单个属性(最上层)
     *
     * @param declaredField 属性
     */
    private void parseField(Field declaredField) {
        if (!declaredField.isAccessible()) {
            // 获取访问权限
            declaredField.setAccessible(true);
        }
        if (declaredField.isAnnotationPresent(CellName.class)) {
            // 判断当前类是否是 Collection 的子类
            dealAttributes(declaredField, false, false);
        } else {
            dealClass(declaredField);
        }
    }

    /**
     * 处理类
     *
     * @param declaredField 属性
     */
    private void dealClass(Field declaredField) {
        Class<?> fieldClass = declaredField.getType();
        // 判断当前类是否是 Collection 的子类
        if (!Collection.class.isAssignableFrom(fieldClass)) {
            throw new AnnotationException(
                "对于非集合类型的字段, 应该使用 @CellName 而不是 @MultipleCell");
        }
        dealMultipleCell(declaredField, fieldClass);
        MultipleCell multipleCell = declaredField.getAnnotation(MultipleCell.class);
        if (multipleCell.basicType()) {
            CellConfig cellConfig =
                new CellConfig(declaredField.getAnnotation(MultipleCell.class).name(),
                    declaredField, fieldClass, true, false);
            sheetConfig.getCellConfigList().add(cellConfig);
        } else {
            // 只支持一层解析
            Class<?> clazz = multipleCell.clazz();
            // 解析实际类型, 并添加到 cellConfigList 中
            for (Field innerDeclaredField : clazz.getDeclaredFields()) {
                // 跳过未声明CellName或者MultipleCell注解的属性
                if (!innerDeclaredField.isAnnotationPresent(CellName.class) ) {
                    continue;
                }
                if (!innerDeclaredField.isAccessible()) {
                    // 获取访问权限
                    innerDeclaredField.setAccessible(true);
                }
                dealAttributes(innerDeclaredField, true, true);
            }
        }
    }

    /**
     * 处理一对多的属性
     * <p>用于存储一对多的属性, 以便于写excel的时候快速获取行号</p>
     *
     * @param declaredField 获取属性名的对象
     * @param fieldClass    属性类
     */
    private void dealMultipleCell(Field declaredField, Class<?> fieldClass) {
        if (sheetConfig.getMultipleCell() != null) {
            throw new AnnotationException(
                "一个 Sheet对象 仅允许使用一个 @MultipleCell 注解");
        }
        CellConfig multipleCellConfig = new CellConfig(declaredField, fieldClass);
        sheetConfig.setMultipleCell(multipleCellConfig);
    }

    /**
     * 处理属性
     *
     * @param declaredField 属性
     */
    private void dealAttributes(Field declaredField, boolean more, boolean nestAttr) {
        Class<?> fieldClass = declaredField.getType();
        if (fieldClass.isAssignableFrom(List.class)) {
            throw new AnnotationException(
                "对于集合类型的字段, 应该使用 @MultipleCell 而不是 @CellName");
        }
        CellConfig cellConfig =
            new CellConfig(declaredField.getAnnotation(CellName.class).value(),
                declaredField, fieldClass, more, nestAttr);
        sheetConfig.getCellConfigList().add(cellConfig);
    }
}
