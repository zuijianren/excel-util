package com.zuijianren.utils.excel.cache;

import com.zuijianren.utils.excel.config.SheetConfig;
import com.zuijianren.utils.excel.core.ConfigParse;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author: 姜辞旧
 * @since: 2022/3/18 9:31
 * @version: 1.0
 */
public class SheetConfigCache {

    private static final Map<Class<?>, SheetConfig> PARSE_MAP = new HashMap<>();

    public static <T> SheetConfig getSheetConfig(Class<T> targetClass) {
        SheetConfig sheetConfig = PARSE_MAP.get(targetClass);
        if (sheetConfig == null) {
            // 解析注解
            sheetConfig = new ConfigParse().parseAnnotation(targetClass);
            PARSE_MAP.put(targetClass, sheetConfig);
        }
        return sheetConfig;
    }

}
