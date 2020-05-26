package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.IExcelImportConsumer;
import cn.com.quanyou.excel.context.impl.ColumnMeta;
import cn.com.quanyou.excel.context.impl.RowData;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.context.ApplicationContext;
import org.springframework.util.ReflectionUtils;

import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

@Slf4j
@Data
@Builder
public class ExcelXlsReader implements HSSFListener {

    private IExcelContext context;
    private ArrayList boundSheetRecords;

    private int sheetIdx = -1;
    private String sheetName;
    private SSTRecord sstRecord;

    private BoundSheetRecord[] orderedBSRs;
    private SheetMeta sheetMeta;
    private RowData rowData;

    private static final int MAX_ROWS_NUM = 1000;

    private FormatTrackingHSSFListener formatListener;
    private ConsumerMeta consumerMeta;

    public void process(POIFSFileSystem fs, ConsumerMeta consumerMeta) throws IOException {

        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener =
                new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
        request.addListenerForAllRecords(workbookBuildingListener);
        factory.processWorkbookEvents(request, fs);
        this.consumerMeta = consumerMeta;
    }


    @Override
    public void processRecord(Record record) {
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add(record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                setSheetMeta(br);
                break;
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;
                setRowData(sheetMeta, brec.getRow(), brec.getColumn(), "");
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;
                setRowData(sheetMeta, berec.getRow(), berec.getColumn(), berec.getBooleanValue());
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;
                setRowData(sheetMeta, lrec.getRow(), lrec.getColumn(), lrec.getValue());
                break;
            case LabelSSTRecord.sid: // 单元格为字符串类型
                LabelSSTRecord lsrec = (LabelSSTRecord) record;
                String value = "";
                if (sstRecord != null)
                    value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
                setRowData(sheetMeta, lsrec.getRow(), lsrec.getColumn(), value);
                break;
            case NumberRecord.sid: // 单元格为数字类型
                NumberRecord numrec = (NumberRecord) record;
                setRowData(sheetMeta, numrec.getRow(), numrec.getColumn(), formatListener.formatNumberDateCell(numrec));
                break;
            case EOFRecord.sid: //Sheet页读取结束时调用
                if (null != sheetMeta) {
                    List<RowData> rowDatas = sheetMeta.getRowDataList();
                    consumerMeta.callUseConsumer(context, sheetMeta, rowDatas);
                }
                break;
            default: //根据配置分批处理
                if (record instanceof LastCellOfRowDummyRecord) {
                    LastCellOfRowDummyRecord rowDummyRecord = (LastCellOfRowDummyRecord) record;

                    //设置一个保险值
                    Integer rowsNum = context.getBatchRows();
                    if (null == rowsNum) {
                        rowsNum = MAX_ROWS_NUM;
                    }

                    if(rowDummyRecord.getRow() >= sheetMeta.getDataRow() - 1) {
                        List<RowData> rowDatas = sheetMeta.getRowDataList();
                        rowDatas.add(rowData);
                        sheetMeta.setRowDataList(rowDatas);
                        if (rowDatas.size() >= rowsNum) {
                            consumerMeta.callUseConsumer(context, sheetMeta, rowDatas);
                        }
                        rowData = new RowData();
                    }
                }
                break;
        }
    }

    private void setSheetMeta(BOFRecord br) {
        if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
            sheetIdx ++;

            if (orderedBSRs == null) {
                orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
            }

            sheetName = orderedBSRs[sheetIdx].getSheetname();
            Optional<SheetMeta> sheetMetaOptional = context.getSheetMetas().stream()
                    .filter(item -> sheetName.equals(item.getSheetName())).findAny();
            sheetMeta = sheetMetaOptional.orElseThrow(() -> new RuntimeException("没有找到合适的Sheet配置文件"));
        }
    }


    private void setRowData(SheetMeta sheetMeta, int thisRow, int thisCol, Object value) {
        if(thisRow >= sheetMeta.getDataRow() - 1) {
            List<ColumnMeta> columnMetas = sheetMeta.getColumnMetas();
            Optional<ColumnMeta> metaOptional =
                    columnMetas.stream().filter(item -> item.getColIdx() == thisCol).findAny();
            ColumnMeta columnMeta = metaOptional.orElseThrow(() -> new RuntimeException("没有找到合适的Column配置文件"));

            Map<String, Object> columnIdxMap = rowData.getColumnToData();
            columnIdxMap.put(columnMeta.getFieldName(), value);

        }
    }
}
