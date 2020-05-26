package cn.com.quanyou.excel.exported.impl;

import cn.com.quanyou.excel.context.impl.ColumnMeta;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import cn.com.quanyou.excel.context.utils.EntityUtils;
import cn.com.quanyou.excel.entity.CellStyleVO;
import cn.com.quanyou.excel.exception.ExcelException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.util.Comparator;
import java.util.Date;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
@Component
public class ExcelColumnHandler {

    /**
     * 根据列生成不同的单元格样式
     *
     * @param wb
     */
    public Function<ColumnMeta, CellStyleVO> createCellStyle(Workbook wb) {

        return columnMeta -> {
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle = setCellFrame(cellStyle);
            if (StringUtils.isNotEmpty(columnMeta.getFormat())) {
                DataFormat format = wb.createDataFormat();
                cellStyle.setDataFormat(format.getFormat(columnMeta.getFormat()));
            }

            CellStyleVO.CellStyleVOBuilder builder = CellStyleVO.builder();
            builder.fieldName(columnMeta.getFieldName());
            builder.cellStyle(cellStyle);

            return builder.build();
        };
    }

    /**
     * 设置边框
     *
     * @param cellStyle
     * @return
     */
    public CellStyle setCellFrame(CellStyle cellStyle) {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        return cellStyle;
    }

    /**
     * 设置单行列数据
     *
     * @param row
     * @param rowIdx
     * @param data
     * @param cellStyleMap
     * @return
     */
    public Consumer<ColumnMeta> setExcelCellData(Row row, int rowIdx, Object data, Map<String, CellStyle> cellStyleMap) {
        return columnMeta -> {
            try {
                int colIdx = columnMeta.getColIdx() - 1;
                Cell cell = row.getCell(colIdx);
                if (null == cell) {
                    cell = row.createCell(colIdx);
                }

                columnMeta.getCellValueLength().add(setExcelCellValue(cell, columnMeta, data, rowIdx, cellStyleMap));
            } catch (Exception ex) {
                log.error(ex.getMessage(), ex);
            }
        };
    }

    /**
     * 自动调整列宽
     *
     * @param sheetMeta
     * @param sheet
     * @return
     */
    public Consumer<ColumnMeta> autoColumnLength(SheetMeta sheetMeta, Sheet sheet) {
        return columnMeta -> {
            int colIdx = columnMeta.getColIdx() - 1;
            if (columnMeta.getWidth() != null && columnMeta.getWidth() > 0) {
                int width = Integer.valueOf(columnMeta.getWidth());
                sheet.setColumnWidth(colIdx, width * 256);
            } else {
                int startNum = 1;
                if (sheetMeta.getTitleRow() > 2) {
                    startNum = 2;
                }

                Cell cell = sheet.getRow(sheetMeta.getTitleRow() - startNum).getCell(colIdx);
                int chineseSize = getChineseSize(cell.getStringCellValue());
                int otherSize = cell.getStringCellValue().length() - chineseSize;
                int titleLength = otherSize + (chineseSize * 2);
                int maxWidth = columnMeta.getCellValueLength().stream().max(Comparator.comparing(i -> i))
                        .orElseThrow(()-> new ExcelException("计算宽度失败，请检查配置列与Excel列是否一致！"));
                log.info(cell+"长度：" + titleLength + " 值长度：" + maxWidth);
                if (titleLength < maxWidth) {
                    sheet.setColumnWidth(colIdx, (maxWidth + 1) * 256);
                } else {
                    sheet.setColumnWidth(colIdx, (titleLength + 1) * 256);
                }
            }

            log.info("列" + columnMeta.getFieldName() + "自动宽度调整为：" + sheet.getColumnWidth(colIdx));
        };
    }


    /**
     * 数据类型设置
     *
     * @param cell
     * @param columnMeta
     * @param obj
     * @param rowIdx
     * @return
     * @throws IllegalAccessException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     */
    private int setExcelCellValue(Cell cell, ColumnMeta columnMeta, Object obj, int rowIdx, Map<String, CellStyle> cellStyleMap)
            throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {

        int length = 0;
        if (StringUtils.isNotEmpty(columnMeta.getFieldName())) {
            if ("SYS_SN".equals(columnMeta.getFieldName())) {
                cell.setCellValue(rowIdx);
                length = String.valueOf(rowIdx).getBytes().length;
            } else {
                if ("double".equals(columnMeta.getType()) || "float".equals(columnMeta.getType())) {
                    Object cellValue = EntityUtils.getProperty(obj, columnMeta.getFieldName());
                    DecimalFormat df = new DecimalFormat("0.######");
                    if (cellValue instanceof Double) {
                        cell.setCellValue(Double.valueOf(df.format(cellValue)));
                    } else if (cellValue instanceof Float) {
                        cell.setCellValue(Float.valueOf(df.format(cellValue)));
                    }

                    length = df.format(cellValue).getBytes().length;
                } else if ("integer".equals(columnMeta.getType())) {
                    Integer cellValue = EntityUtils.getProperty(obj, columnMeta.getFieldName());
                    cell.setCellValue(cellValue);
                    length = String.valueOf(cellValue).getBytes().length;

                } else if ("date".equals(columnMeta.getType())) {
                    Date cellValue = EntityUtils.getProperty(obj, columnMeta.getFieldName());
                    cell.setCellValue(cellValue);
                    length = String.valueOf(cellValue).getBytes().length;

                } else if ("rate".equals(columnMeta.getType())) {
                    Float cellValue = EntityUtils.getProperty(obj, columnMeta.getFieldName());
                    cell.setCellValue(cellValue);
                    length = String.valueOf(cellValue).getBytes().length;
                } else {
                    String cellValue = BeanUtils.getProperty(obj, columnMeta.getFieldName());
                    cell.setCellValue(cellValue);

                    if (StringUtils.isNotEmpty(cellValue)) {
                        int chineseSize = getChineseSize(cellValue);
                        int otherSize = cellValue.length() - chineseSize;
                        length = otherSize + (chineseSize * 2); //中文字符按2个字节算长度
                    }
                }
            }
        }

        //单元格为空是
        if (cell == null || "".equals(cell)) {
            String cellValue = columnMeta.getDefaultValue();
            if (StringUtils.isNotEmpty(cellValue)) {
                cell.setCellValue(cellValue);
                length = cellValue.getBytes().length;
            }
        }

        cell.setCellStyle(cellStyleMap.get(columnMeta.getFieldName()));
        return length;
    }

    /**
     * 计算中文字符长度
     *
     * @param content
     * @return
     */
    private int getChineseSize(String content) {
        int count = 0;//汉字数量
        String regEx = "[\\u4e00-\\u9fa5]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(content);
        int len = m.groupCount();
        while (m.find()) {
            for (int i = 0; i <= len; i++) {
                count = count + 1;
            }
        }
        return count;
    }
}
