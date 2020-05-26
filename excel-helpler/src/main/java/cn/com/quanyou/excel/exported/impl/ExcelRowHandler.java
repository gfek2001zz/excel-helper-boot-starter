package cn.com.quanyou.excel.exported.impl;

import cn.com.quanyou.excel.context.impl.ColumnMeta;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

@Component
public class ExcelRowHandler {

    @Autowired
    private ExcelColumnHandler excelColumnHandler;


    /**
     * 创建标题行
     *
     * @param columnMetas
     * @param sheetMeta
     * @param sheet
     */
    public void createTitleRow(List<ColumnMeta> columnMetas, SheetMeta sheetMeta, Sheet sheet) {
        Workbook wb = sheet.getWorkbook();

        Row row = sheet.createRow(sheetMeta.getTitleRow() - 1);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle = excelColumnHandler.setCellFrame(cellStyle);

        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直
        cellStyle.setAlignment(HorizontalAlignment.CENTER);// 水平

        Font font = wb.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFont(font);
        int cellIdx = 0;
        for (ColumnMeta columnMeta : columnMetas) {
            Cell cell = row.createCell(cellIdx);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(columnMeta.getDisplayName());

            cellIdx ++;
        }
    }

    /**
     * 创建数据行
     *
     * @param dataList
     * @param dataRow
     * @param columnMetas
     * @param sheet
     * @param rowIdx
     * @param cellStyleMap
     */
    public int createDataRow(List<?> dataList, int dataRow, List<ColumnMeta> columnMetas, Sheet sheet, int rowIdx, Map<String, CellStyle> cellStyleMap) {
        if(rowIdx == 0) {
            rowIdx = dataRow - 1;
        }

        Integer rowNum = 1;

        for (Object data : dataList) {
            Row row = sheet.getRow(rowIdx);
            if (null == row) {
                row = sheet.createRow(rowIdx);
            }

            columnMetas.stream().forEach(excelColumnHandler.setExcelCellData(row, rowNum, data, cellStyleMap));

            rowNum ++;
            rowIdx ++;
        }

        return rowIdx;
    }
}
