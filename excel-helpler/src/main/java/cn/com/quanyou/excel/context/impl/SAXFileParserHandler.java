package cn.com.quanyou.excel.context.impl;

import cn.com.quanyou.excel.context.utils.EntityUtils;
import lombok.extern.slf4j.Slf4j;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public class SAXFileParserHandler extends DefaultHandler {
    String value = null;

    private final ExcelContext.ExcelContextBuilder contextBuilder;
    private SheetMeta.SheetMetaBuilder sheetMetaBuilder;
    private ColumnMeta.ColumnMetaBuilder columnMetaBuilder;

    private Integer colIdx = 0;

    public SAXFileParserHandler(ExcelContext.ExcelContextBuilder contextBuilder) {
        this.contextBuilder = contextBuilder;
    }

    List<SheetMeta> sheetMetas = new ArrayList<>();
    public List<SheetMeta> getSheetMetas() {
        return sheetMetas;
    }

    @Override
    public void startDocument() throws SAXException {
        super.startDocument();
        log.debug("SAX解析开始");
    }

    @Override
    public void endDocument() throws SAXException {
        super.endDocument();
        log.debug("SAX解析结束");
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        super.startElement(uri, localName, qName, attributes);

        log.debug("uri:" + uri);
        log.debug("localName:" + localName);
        log.debug("qName:" + qName);

        if ("sheet".equals(qName)) {
            colIdx = 0;
            sheetMetaBuilder = SheetMeta.builder();
            EntityUtils.parseXML2Entity(sheetMetaBuilder.build(), attributes);

        } else if ("column".equals(qName)) {
            columnMetaBuilder = ColumnMeta.builder();
            EntityUtils.parseXML2Entity(columnMetaBuilder.build(), attributes);

            columnMetaBuilder.colIdx(colIdx);
            colIdx = colIdx + 1;
        } else if ("excel".equals(qName)) {
            int attrNum = attributes.getLength();
            for (int i = 0; i < attrNum; i ++) {
                String attrName = attributes.getQName(i);
                if ("name".equals(attrName)) {
                    String fileName = EntityUtils.getExcelFilePath(contextBuilder.build(), attributes.getValue(i));
                    contextBuilder.fileName(fileName);
                } else {
                    EntityUtils.parseXML2Entity(contextBuilder.build(), attributes);
                }
            }
        }
    }


    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        super.endElement(uri, localName, qName);

        if ("sheet".equals(qName)) {
            sheetMetas.add(sheetMetaBuilder.build());
        } else if ("column".equals(qName)) {
            columnMetaBuilder.displayName(value);
            List<ColumnMeta> columnMetas = sheetMetaBuilder.build().getColumnMetas();
            columnMetas.add(columnMetaBuilder.build());
            sheetMetaBuilder.columnMetas(columnMetas);
        }
    }

    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        // TODO Auto-generated method stub
        super.characters(ch, start, length);
        value = new String(ch, start, length);
        if (!value.trim().equals("")) {
            log.debug("节点值是：" + value);
        }
    }
}
