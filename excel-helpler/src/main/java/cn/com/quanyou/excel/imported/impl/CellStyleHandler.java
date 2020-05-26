package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.excel.context.impl.ColumnMeta;
import lombok.Data;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.stereotype.Component;
import org.xml.sax.Attributes;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;

@Component(value = "cellStyleHandler")
public class CellStyleHandler {

    private final DataFormatter formatter = new DataFormatter();


    /**
     * 单元格格式化
     *
     * @param attributes
     * @param stylesTable
     * @return
     */
    public CellFormat getCellFormat(Attributes attributes, StylesTable stylesTable) {
        short formatIndex = -1;
        String formatString = null;
        String cellStyleStr = attributes.getValue("s");

        if (cellStyleStr != null){
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();
            if ("m/d/yy" == formatString){
                formatString = "yyyy-MM-dd";
            }
            if (formatString == null){
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
        }

        CellFormat result = new CellFormat();
        result.setFormatIndex(formatIndex);
        result.setFormatString(formatString);

        return result;
    }

    /**
     * 获取转换后的数据
     *
     * @param columnMeta
     * @param value
     * @return
     */
    public Object getDataValue(ColumnMeta columnMeta, CellFormat cellFormat, String value) {

        String columnType = columnMeta.getType();
        String formatString = cellFormat.getFormatString();
        short formatIndex = cellFormat.getFormatIndex();

        Object result;
        String thisStr;
        switch (columnType) {
            //这几个的顺序不能随便交换，交换了很可能会导致数据错误
            case "boolean":
                char first = value.charAt(0);
                result = first == '0' ? true : false;
                break;
            case "formula":
                result = '"' + value + '"';
                break;
            case "inlinestr":
                XSSFRichTextString rtsi = new XSSFRichTextString(value);
                result = rtsi.toString();
                break;
            case "string":
                result = value;
                break;
            case "short":
                result = new Short(value);
                break;
            case "double":
                if (formatString != null){
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }

                result = new BigDecimal(thisStr.replace("_", "").trim()).doubleValue();
                break;
            case "float":
                if (formatString != null){
                    thisStr = formatter.formatRawCellContents(Float.parseFloat(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }

                result = new BigDecimal(thisStr.replace("_", "").trim()).floatValue();
                break;
            case "integer":
                if (formatString != null){
                    thisStr = formatter.formatRawCellContents(Integer.parseInt(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }

                result = new BigDecimal(thisStr.replace("_", "").trim()).intValue();
                break;
            case "date":
                try{
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                    thisStr = thisStr.replace(" ", "");
                    SimpleDateFormat format = new SimpleDateFormat(formatString);
                    result = format.parse(thisStr);
                }catch(Exception ex){
                    result = value;
                }
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    @Data
    public class CellFormat {
        private short formatIndex;
        private String formatString;
    }
}
