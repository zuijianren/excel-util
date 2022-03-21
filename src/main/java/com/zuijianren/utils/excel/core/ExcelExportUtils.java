package com.zuijianren.utils.excel.core;

import com.zuijianren.utils.excel.cache.SheetConfigCache;
import com.zuijianren.utils.excel.exceptions.WriteToFileException;
import com.zuijianren.utils.excel.pojo.ExportData;
import com.zuijianren.utils.excel.theme.DefaultTheme;
import com.zuijianren.utils.excel.theme.Theme;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.List;
import lombok.extern.java.Log;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author: 姜辞旧
 * @since: 2022/3/16 15:15
 * @version: 1.0
 */
@Log
public class ExcelExportUtils {

    private static final String FILE_SUFFIX = ".xlsx";

    private Theme theme = new DefaultTheme();

    public ExcelExportUtils() {
    }

    public ExcelExportUtils(Theme theme) {
        this.theme = theme;
    }

    /**
     * 导出到 XSSFWorkbook 中
     *
     * @param list  数据
     * @param clazz 数据类型
     * @return XSSFWorkbook 对象
     */
    public <E> XSSFWorkbook export(List<E> list, Class<E> clazz) {
        return export(Collections.singletonList(new ExportData(clazz, list)));
    }

    /**
     * 导出到 XSSFWorkbook 中
     *
     * @param exportData 导出数据
     * @return XSSFWorkbook 对象
     */
    public XSSFWorkbook export(ExportData exportData) {
        return export(Collections.singletonList(exportData));
    }

    /**
     * 导出到 XSSFWorkbook 中
     *
     * @param exportDataList 导出数据
     * @return XSSFWorkbook 对象
     */
    public XSSFWorkbook export(List<ExportData> exportDataList) {
        // 初始化
        XSSFWorkbook workbook = new XSSFWorkbook();
        theme.init(workbook);

        // 遍历写入各个 Sheet 中
        exportDataList.forEach(exportData -> {
            // 创建excel相关类, 实现文件编写
            SheetWriter sheetWriter = new SheetWriter(
                SheetConfigCache.getSheetConfig(exportData.getClazz()), theme);
            sheetWriter.writeToSheet(exportData.getList(), workbook);
        });

        return workbook;
    }

    /**
     * 导出数据到指定文件
     *
     * @param file 文件
     */
    public void exportToFile(ExportData exportData, File file) {
        exportToFile(Collections.singletonList(exportData), file);
    }

    /**
     * 导出数据到指定文件
     *
     * @param file 文件
     */
    public void exportToFile(List<ExportData> exportDataList, File file) {
        validateFileFormat(file);
        XSSFWorkbook workbook = export(exportDataList);
        try {
            workbook.write(new FileOutputStream(file));
        } catch (IOException e) {
            throw new WriteToFileException(e.getMessage());
        }
    }

    /**
     * 校验文件格式
     *
     * @param file 文件
     */
    private void validateFileFormat(File file) {
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (!FILE_SUFFIX.equals(fileType)) {
            throw new IllegalArgumentException("文件格式错误, 文件类型应为 \".xlsx\"");
        }
    }


}
