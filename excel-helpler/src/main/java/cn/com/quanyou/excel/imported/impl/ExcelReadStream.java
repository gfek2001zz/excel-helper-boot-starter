package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.impl.ErrorMeta;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import cn.com.quanyou.excel.dao.IExcelTaskDao;
import cn.com.quanyou.excel.entity.ExcelTypeEnum;
import cn.com.quanyou.excel.entity.TaskModel;
import cn.com.quanyou.excel.entity.TaskStatusEnum;
import cn.com.quanyou.excel.exported.impl.ExcelFileUploadHandler;
import cn.com.quanyou.excel.imported.IExcelReadStream;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

@Slf4j
@Component
public class ExcelReadStream implements IExcelReadStream {

    @Autowired
    private ConsumerMeta consumerMeta;

    @Autowired
    private IExcelTaskDao excelTaskDao;

    @Autowired
    private ExcelFileUploadHandler fileUploadHandler;

    @Async
    public void readExcel(IExcelContext context, ByteArrayOutputStream fileOutputStream) {
        excelTaskDao.updateTaskStatus(TaskStatusEnum.PROCESSING, context.getTaskId());
        InputStream fileInputStream = null;
        InputStream sheetStream = null;
        POIFSFileSystem fs = null;
        OPCPackage pkg = null;

        TaskModel taskModel = excelTaskDao.findExcelTaskByTaskId(context.getTaskId());
        String downloadUrl = fileUploadHandler.uploadExcelFile(context, fileOutputStream);
        taskModel.setDownloadUrl(downloadUrl);

        try {
            fileInputStream = new ByteArrayInputStream(fileOutputStream.toByteArray());
            if (ExcelTypeEnum.XLS.equals(ExcelTypeEnum.valueOf(context.getExcelType()))) {
                //Excel2003大数据导入
                fs = new POIFSFileSystem(fileInputStream);
                ExcelXlsReader.ExcelXlsReaderBuilder readerBuilder = ExcelXlsReader.builder();
                readerBuilder.context(context);
                ExcelXlsReader excelXlsReader = readerBuilder.build();
                excelXlsReader.process(fs, consumerMeta);
            } else if (ExcelTypeEnum.XLSX.equals(ExcelTypeEnum.valueOf(context.getExcelType()))) {

                //Excel2007使用SAX方式大数据导入
                pkg = OPCPackage.open(fileInputStream);
                XSSFReader reader = new XSSFReader(pkg);
                SharedStringsTable sst = reader.getSharedStringsTable();
                List<SheetMeta> sheetMetas = context.getSheetMetas();
                for (SheetMeta sheetMeta : sheetMetas) {
                    XMLReader parser = fetchSheetParser(context, sheetMeta, sst, reader);
                    sheetStream = reader.getSheet(String.format("rId%d", getSheetNum(reader, sheetMeta.getSheetName())));
                    InputSource sheetSource = new InputSource(sheetStream);
                    parser.parse(sheetSource);
                    sheetStream.close();
                }
            }

            taskModel.setStatus(TaskStatusEnum.PROCESSED);
        } catch (Exception ex) {
            context.setThrowable(ex);
            log.error(ex.getMessage(), ex);
        } finally {
            closeStream(fileOutputStream, fileInputStream, sheetStream, fs, pkg);
            String errorMessage = getErrorMessage(context);

            //写入错误消息
            taskModel.setFileName(context.getFileName());


            if (StringUtils.isNotEmpty(errorMessage)) {
                taskModel.setStatus(TaskStatusEnum.ERROR);
                taskModel.setExecResults(errorMessage);
            } else if (context.getThrowable() != null) {
                taskModel.setStatus(TaskStatusEnum.ERROR);
                StringWriter stringWriter = new StringWriter();
                context.getThrowable().printStackTrace(new PrintWriter(stringWriter));
                taskModel.setExecResults(stringWriter.getBuffer().toString());
            }

            taskModel.setEndTime(new Date());
            excelTaskDao.updateTaskInfo(taskModel);
        }
    }

    /**
     * 错误原因打印
     *
     * @param context
     * @return
     */
    private String getErrorMessage(IExcelContext context) {
        List<ErrorMeta> errorMetas = excelTaskDao.findErrorMeta(context.getTaskId());
        context.getErrorMetas().addAll(errorMetas);
        excelTaskDao.deleteErrorMeta(context.getTaskId());

        StringBuffer buffer = new StringBuffer();
        return context.getErrorMetas().stream().map(errorMeta -> {
            buffer.append("行：").append(StringUtils.defaultString(errorMeta.getRowName(), ""))
                    .append("，列：").append(StringUtils.defaultString(errorMeta.getColName(), ""))
                    .append("，原因：").append(errorMeta.getMessage());

            String result = buffer.toString();
            buffer.setLength(0);

            return result;
        }).collect(Collectors.joining("<br />"));
    }


    /**
     * 关闭流
     *
     * @param out
     * @param in
     * @param sheetStream
     * @param fs
     * @param pkg
     */
    private void closeStream(ByteArrayOutputStream out, InputStream in, InputStream sheetStream, POIFSFileSystem fs, OPCPackage pkg) {
        if (out != null) {
            try {
                out.flush();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
            try {
                out.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }

        if (in != null) {
            try {
                in.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }

        if (sheetStream != null) {
            try {
                sheetStream.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }

        if (pkg != null) {
            try {
                pkg.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }

        if (fs != null) {
            try {
                fs.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
    }

    /**
     * 查找当前指定Sheet页位置
     *
     * @param reader
     * @param sheetName
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    private int getSheetNum(XSSFReader reader, String sheetName) throws IOException, InvalidFormatException {
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
        int sheetNum = 0;

        if (sheetName != null) {
            while (sheets.hasNext()) {
                sheetNum++;
                InputStream sheet = sheets.next();
                if (sheets.getSheetName().equals(sheetName)) {
                    sheet.close();
                    break;
                }

                sheet.close();
            }
        } else {
            sheetNum = 1;
        }

        return sheetNum;
    }

    public XMLReader fetchSheetParser(IExcelContext context, SheetMeta sheetMeta, SharedStringsTable sst, XSSFReader reader) throws Exception {
        XMLReader parser =
                XMLReaderFactory.createXMLReader(
                        "org.apache.xerces.parsers.SAXParser"
                );
        ContentHandler handler = new ExcelXlsxReader(context, sheetMeta, sst, reader);
        parser.setContentHandler(handler);
        return parser;
    }
}
