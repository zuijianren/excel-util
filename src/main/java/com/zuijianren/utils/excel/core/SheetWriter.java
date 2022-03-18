package com.zuijianren.utils.excel.core;

import com.zuijianren.utils.excel.config.CellConfig;
import com.zuijianren.utils.excel.config.SheetConfig;
import com.zuijianren.utils.excel.exceptions.WriteToFileException;
import com.zuijianren.utils.excel.theme.Theme;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 实际执行 excel 写入的类
 *
 * @author: 姜辞旧
 * @since: 2022/3/17 14:28
 * @version: 1.0
 */
public class SheetWriter {

    private final SheetConfig sheetConfig;

    private final Theme theme;

    /**
     * 起始列
     */
    private final int initNum;

    public SheetWriter(SheetConfig sheetConfig, Theme theme) {
        this.sheetConfig = sheetConfig;
        this.theme = theme;
        initNum = theme.isShowNum() ? 1 : 0;
    }

    /**
     * 将数据写入到 sheet
     *
     * @param list     数据
     * @param workbook 文本薄
     */
    public void writeToSheet(List<?> list, XSSFWorkbook workbook) {
        XSSFSheet sheet = workbook.createSheet(sheetConfig.getSheetName());
        // 行号
        int rowNum = 0;
        // 创建 top
        if (theme.isCreateTop()) {
            createTop(sheet, rowNum++);
        }
        // 创建 title
        createTitle(sheet, rowNum++);
        // 创建 content
        try {
            createContent(sheet, list, rowNum);
        } catch (IllegalAccessException e) {
            throw new WriteToFileException("写入文件错误" + e.getMessage());
        }
    }

    /**
     * 创建顶行
     *
     * @param sheet  表格
     * @param rowNum 行号
     */
    private void createTop(XSSFSheet sheet, int rowNum) {
        XSSFRow titleRow = sheet.createRow(rowNum);
        int size = sheetConfig.getCellConfigList().size();
        // 如果需要展示序号, 则扩充一格
        if (theme.isShowNum()) {
            size++;
        }
        mergeColCell(sheet, rowNum, 0, size);
        XSSFCell cell = titleRow.createCell(0);
        cell.setCellStyle(theme.getTopStyle());
        cell.setCellValue(sheetConfig.getSheetName());
    }

    /**
     * 创建标题行
     *
     * @param sheet  表格
     * @param rowNum 行号
     */
    private void createTitle(XSSFSheet sheet, int rowNum) {
        XSSFRow titleRow = sheet.createRow(rowNum);
        List<CellConfig> cellConfigList = sheetConfig.getCellConfigList();
        if (initNum == 1) {
            XSSFCell numberCell = titleRow.createCell(0);
            numberCell.setCellStyle(theme.getTitleStyle());
            numberCell.setCellValue("序号");
        }
        // 根据是否有序列号进行移动
        for (int i = initNum; i < cellConfigList.size() + initNum; i++) {
            XSSFCell cell = titleRow.createCell(i);
            String cellName = cellConfigList.get(i - initNum).getCellName();
            cell.setCellStyle(theme.getTitleStyle());
            cell.setCellValue(cellName);
        }
    }

    /**
     * 创建内容
     * <p>todo 待完善 一对多的情况</p>
     *
     * @param sheet sheet页
     * @param list  对应数据
     * @throws IllegalAccessException 反射获取值错误
     */
    private void createContent(XSSFSheet sheet, List<?> list, int rowNum)
        throws IllegalAccessException {
        List<CellConfig> cellConfigList = sheetConfig.getCellConfigList();
        for (int i = 0; i < list.size(); i++) {
            // 判断是否多行, 多行走指定循环
            if (sheetConfig.getMultipleCell() == null) {
                writeOne(sheet, list.get(i), cellConfigList, rowNum, i);
                // 行数加1
                rowNum++;
            } else {
                int size = writeMore(sheet, list, rowNum, cellConfigList, i);
                // 行数加指定范围
                rowNum += size;
            }
        }

    }

    private int writeMore(XSSFSheet sheet, List<?> list, int rowNum,
        List<CellConfig> cellConfigList, int index) throws IllegalAccessException {
        // 获取容器元素
        List obj = (List) sheetConfig.getMultipleCell().getField()
            .get(list.get(index));
        // 获取大小
        int size = obj.size();
        // 需要合并单元格的属性
        List<CellConfig> mergeCellList = new ArrayList<>();
        // 遍历
        for (int j = 0; j < size; j++) {
            XSSFRow row = sheet.createRow(rowNum + j);
            // 第一行数据 && 需要写入序号
            if (j == 0 && initNum == 1) {
                // 写入序号
                writeValue(row, 0, index + 1);
            }
            for (int column = initNum; column < cellConfigList.size() + initNum; column++) {
                CellConfig cellConfig = cellConfigList.get(column - initNum);
                if (cellConfig.isMore()) {
                    if (cellConfig.isNestAttr()) {
                        // 嵌套属性从target获取
                        writeValue(row, column, cellConfig.getField().get(obj.get(j)));
                    } else {
                        writeValue(row, column, obj.get(j));
                    }
                } else {
                    // 其余属性, 从最外层对象获取
                    writeValue(row, column, cellConfig.getField().get(list.get(index)));
                    if (!mergeCellList.contains(cellConfig)) {
                        mergeCellList.add(cellConfig);
                    }
                }
            }
        }
        if (initNum == 1) {
            // 如果有序列号, 合并序列行
            mergeRowCell(sheet, rowNum, 0, size);
        }
        // 合并单元格
        mergeCellList.forEach(e -> {
            mergeRowCell(sheet, rowNum, cellConfigList.indexOf(e) + initNum, size);
        });
        return size;
    }

    /**
     * 写入单条数据
     *
     * @param sheet          sheet表
     * @param obj            数据
     * @param cellConfigList 配置
     * @param rowNum         行
     * @param index          序号
     * @throws IllegalAccessException 反射异常
     */
    private void writeOne(XSSFSheet sheet, Object obj, List<CellConfig> cellConfigList, int rowNum,
        int index)
        throws IllegalAccessException {
        XSSFRow row = sheet.createRow(rowNum);
        if (initNum == 1) {
            // 写入序号
            writeValue(row, 0, index + 1);
        }
        // 遍历字段, 填入表格
        for (int column = initNum; column < cellConfigList.size() + initNum; column++) {
            writeValue(row, column, cellConfigList.get(column - initNum).getField().get(obj));
        }
    }

    /**
     * 写入值
     *
     * @param row    行
     * @param column 列
     * @param value  值
     */
    private void writeValue(XSSFRow row, int column, Object value) {
        XSSFCell cell = row.createCell(column);
        cell.setCellStyle(theme.getContentStyle());
        cell.setCellValue(value.toString());
    }


    /**
     * 横向合并单元格
     *
     * @param sheet  sheet对象
     * @param rowNum 起始行
     * @param colNum 起始列
     * @param size   大小
     */
    private void mergeRowCell(XSSFSheet sheet, int rowNum, int colNum, int size) {
        // 跳过小于2的情况
        if (size <= 1) {
            return;
        }
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum + size - 1, colNum, colNum));
    }


    /**
     * 纵向合并单元格
     *
     * @param sheet  sheet对象
     * @param colNum 起始列
     * @param rowNum 起始行
     * @param size   大小
     */
    private void mergeColCell(XSSFSheet sheet, int rowNum, int colNum, int size) {
        // 跳过小于2的情况
        if (size <= 1) {
            return;
        }
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, colNum, colNum + size - 1));
    }
}
