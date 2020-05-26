package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.common.entity.BaseModel;
import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.IExcelImportConsumer;
import cn.com.quanyou.excel.context.impl.ErrorMeta;
import cn.com.quanyou.excel.context.impl.ExcelContext;
import cn.com.quanyou.excel.context.impl.RowData;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import cn.com.quanyou.excel.exception.ExcelException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Component;
import org.springframework.util.ReflectionUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
@Component(value = "consumerMeta")
public class ConsumerMeta {


    public void callBegin(IExcelContext context, SheetMeta sheetMeta) {
        IExcelImportConsumer consumer = getExcelImportConsumer(context, sheetMeta.getConsumerBean());
        consumer.begin(context);
    }

    /**
     * 调用消费类消费数据
     *
     * @param context
     * @param sheetMeta
     * @param rowDatas
     */
    public void callUseConsumer(IExcelContext context, SheetMeta sheetMeta, List<RowData> rowDatas) {
        IExcelImportConsumer consumer = getExcelImportConsumer(context, sheetMeta.getConsumerBean());

        try {
            List<Object> data = new ArrayList<>();
            for (RowData rowData : rowDatas) {
                Object voObj = getVoBeanInstance(sheetMeta.getVoBean());
                Map<String, Object> columnData = rowData.getColumnToData();
                List<ErrorMeta> errorMetas = context.getErrorMetas();

                //有错误则不消费这一行数据
                if (CollectionUtils.isEmpty(errorMetas) ||
                        errorMetas.stream().anyMatch(item -> columnData.get(item.getFieldName()) == null)) {

                    if (voObj instanceof BaseModel) {
                        columnData.keySet().stream().forEach(key -> {
                            try {
                                Field field = ReflectionUtils.findField(voObj.getClass(), key);
                                if (null != field) {
                                    field.setAccessible(true);
                                    if (columnData.get(key) != null) {
                                        field.set(voObj, columnData.get(key));
                                    }
                                } else {
                                    throw new ExcelException(key.concat("无法找到对应的实体字段"));
                                }
                            } catch (Exception ex) {
                                throw new ExcelException(ex);
                            }
                        });

                        ((BaseModel) voObj).setRowIdx(rowData.getRowIdx());
                        ((BaseModel) voObj).setColumnNameMap(rowData.getColumnIdxToColumn());
                        data.add(voObj);
                    } else {
                        throw new ExcelException(sheetMeta.getVoBean().concat("实体需要继承BaseModel"));
                    }
                }
            }
            consumer.useBatchData(context, data);
            rowDatas.clear();
        } catch (Exception ex) {
            log.error(ex.getMessage(), ex);
            consumer.fail(context);
            throw new ExcelException(ex);
        }
    }


    public void callEnd(IExcelContext context, SheetMeta sheetMeta) {
        IExcelImportConsumer consumer = getExcelImportConsumer(context, sheetMeta.getConsumerBean());
        consumer.end(context);
    }

    private Object getVoBeanInstance(String voBean) {
        Object result;

        try {
            Class<?> clazz = Class.forName(voBean);
            result = clazz.newInstance();
        } catch (Exception ex) {
            log.error(ex.getMessage(), ex);
            throw new ExcelException("sheet页映射的vo类配置错误！" + ex);
        }

        return result;
    }


    private IExcelImportConsumer getExcelImportConsumer(IExcelContext context, String consumerBean) {
        ApplicationContext applicationContext = context.getApplicationContext();
        return applicationContext.getBean(consumerBean, IExcelImportConsumer.class);
    }
}
