package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.impl.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Optional;

@Slf4j
public class ExcelXlsxReader extends DefaultHandler {

    private SharedStringsTable sst;
    private int curRow = 1;
    private int curCol = 1;

    private String lastContents;
    private boolean nextIsString;

    //定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
    private String preRef = null;
    private String ref = null;

    private RowData rowData;
    private SheetMeta sheetMeta;
    private IExcelContext context;
    boolean nullFlag = true;

    private StylesTable stylesTable;
    private CellStyleHandler cellStyleHandler;
    private ConsumerMeta consumerMeta;
    private ValidatorMeta validatorMeta;

    private CellStyleHandler.CellFormat cellFormat;

    private static final int MAX_ROWS_NUM = 1000;

    public ExcelXlsxReader(IExcelContext context, SheetMeta sheetMeta, SharedStringsTable sst, XSSFReader reader)
            throws IOException, InvalidFormatException {
        this.sst = sst;
        this.context = context;
        this.sheetMeta = sheetMeta;
        this.stylesTable = reader.getStylesTable();

        this.cellStyleHandler = context.getApplicationContext().getBean("cellStyleHandler", CellStyleHandler.class);
        this.consumerMeta = context.getApplicationContext().getBean("consumerMeta", ConsumerMeta.class);
        this.validatorMeta = context.getApplicationContext().getBean("validatorMeta", ValidatorMeta.class);
    }

    @Override
    public void startDocument() {
        log.info("--读取excel开始");
        ((ExcelContext) context).setCurrentSheet(sheetMeta);
        consumerMeta.callBegin(context, sheetMeta);
    }

    /**
     * 开始读取单元格时触发
     *
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        if ("c".equals(name)) {
            nullFlag = true;
            //前一个单元格的位置
            if(preRef == null) {
                preRef = attributes.getValue("r");
            } else {
                preRef = ref;
            }

            //当前单元格的位置
            ref = attributes.getValue("r");
            cellFormat = cellStyleHandler.getCellFormat(attributes, stylesTable);

            String cellType = attributes.getValue("t");
            if(cellType != null && cellType.equals("s")) {
                nextIsString = true;
            } else {
                nextIsString = false;
            }

            // Clear contents cache
            lastContents = "";
        }
    }

    /**
     * 单元格读取结束时触发
     *
     * @param uri
     * @param localName
     * @param name
     */
    @Override
    public void endElement(String uri, String localName, String name) {

        if(nextIsString) {
            int idx = Integer.parseInt(lastContents);
            lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            nextIsString = false;
        }

        if("v".equals(name)) {
            //获取每个单元格的值
            log.info("单元格：{}", lastContents);
            nullFlag = false;
            setRowData(lastContents.trim());

            curCol ++;
        } else if (nullFlag && "c".equals(name)) {
            setRowData("");
            curCol ++;
        } else if ("row".equals(name)) {
            //设置一个保险值
            Integer rowsNum = context.getBatchRows();
            if (null == rowsNum) {
                rowsNum = MAX_ROWS_NUM;
            }

            List<RowData> rowDatas = sheetMeta.getRowDataList();
            if (rowData != null) {
                if (MapUtils.isEmpty(context.getCurrentSheetColumnMap())) {
                    ((ExcelContext) context).setCurrentSheetColumnMap(rowData.getColumnIdxToColumn());
                }

                rowDatas.add(rowData);
                sheetMeta.setRowDataList(rowDatas);
                if (rowDatas.size() >= rowsNum) {
                    consumerMeta.callUseConsumer(context, sheetMeta, rowDatas);
                }
            }

            rowData = new RowData();
            curCol = 1;
            preRef = null;
            ref = null;

            curRow ++;
        }
    }

    /**
     * Sheet读取结束触发
     *
     */
    @Override
    public void endDocument() {
        log.info("--读取excel结束");
        List<RowData> rowDatas = sheetMeta.getRowDataList();
        if (CollectionUtils.isNotEmpty(rowDatas)) {
            consumerMeta.callUseConsumer(context, sheetMeta, rowDatas);
        }

        consumerMeta.callEnd(context, sheetMeta);
    }

    /**
     * 获取单元格的值
     *
     * @param ch
     * @param start
     * @param length
     * @throws SAXException
     */
    public void characters(char[] ch, int start, int length) {
        lastContents += new String(ch, start, length);
    }

    /**
     * 获取行号
     *
     * @return
     */
    private int getColumnIdx(String ref, String compareRef) {
        String refXfd = ref.replaceAll("\\d+", "");
        String compareXfd = compareRef.replaceAll("\\d+", "");

        refXfd = fillChar(refXfd, 3, '@');
        compareXfd = fillChar(compareXfd, 3, '@');

        char[] letter = refXfd.toCharArray();
        char[] compareLetter = compareXfd.toCharArray();
        int res = (letter[0] - compareLetter[0]) * 26 * 26 + (letter[1] - compareLetter[1]) * 26 + (letter[2] - compareLetter[2]);

        return res;
    }


    /**
     * 字符串的填充
     * @param str
     * @param len
     * @param let
     * @return
     */
    String fillChar(String str, int len, char let){
        int strLen = str.length();
        if(strLen < len){
            for(int i=0; i<(len- strLen); i++){
                str = let + str;
            }
        }
        return str;
    }

    /**
     * 构建行数据
     *
     * @param value
     */
    private void setRowData(String value) {
        if(curRow >= sheetMeta.getDataRow()) {
            List<ColumnMeta> columnMetas = sheetMeta.getColumnMetas();
            Optional<ColumnMeta> metaOptional =
                    columnMetas.stream().filter(item -> item.getColIdx() == getColumnIdx(ref, "A") + 1).findAny();

            if (metaOptional.isPresent()) {
                ColumnMeta columnMeta = metaOptional.get();
                validatorMeta.callValidator(context, columnMeta, ref, value);

                Map<String, Object> columnIdxMap = rowData.getColumnToData();
                columnIdxMap.put(columnMeta.getFieldName(), cellStyleHandler.getDataValue(columnMeta, cellFormat, value));
                rowData.setColumnToData(columnIdxMap);

                Map<String, String> columnIdxToColumnMap = rowData.getColumnIdxToColumn();
                columnIdxToColumnMap.put(columnMeta.getFieldName(), ref.replaceAll("\\d+", ""));
                rowData.setColumnIdxToColumn(columnIdxToColumnMap);

                rowData.setRowIdx(curRow);
            }
        }
    }
}
