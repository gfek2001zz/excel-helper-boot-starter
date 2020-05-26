package cn.com.quanyou.excel.context.impl;

import lombok.extern.slf4j.Slf4j;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ExcelXmlParser {

    public List<SheetMeta> parse(ExcelContext.ExcelContextBuilder contextBuilder) {
        List<SheetMeta> sheetMetas = new ArrayList<>();
        SAXParserFactory factory = SAXParserFactory.newInstance();

        try {
            SAXFileParserHandler handler = new SAXFileParserHandler(contextBuilder);
            SAXParser parser = factory.newSAXParser();

            parser.parse(contextBuilder.build().getXmlFileStream(), handler);
            sheetMetas = handler.getSheetMetas();

        }  catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            return sheetMetas;
        }
    }

}
