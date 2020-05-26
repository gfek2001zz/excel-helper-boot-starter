package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.IExcelImportValidator;
import cn.com.quanyou.excel.context.impl.ColumnMeta;
import cn.com.quanyou.excel.context.impl.ErrorMeta;
import cn.com.quanyou.excel.entity.ValidatorResult;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Component;

import java.util.List;

@Slf4j
@Component(value = "validatorMeta")
public class ValidatorMeta {

    /**
     * 调用校验器
     *
     * @param context
     * @param cellName
     * @param cellValue
     */
    public void callValidator(IExcelContext context, ColumnMeta columnMeta, String cellName, Object cellValue) {
        List<String> validatorBeans = columnMeta.getValidatorBean();

        if (CollectionUtils.isNotEmpty(validatorBeans)) {
            validatorBeans.stream().forEach(validatorBean -> {
                IExcelImportValidator validator = getExcelImportValidator(context, validatorBean);
                if (null != validator) {
                    ValidatorResult result = validator.validate(columnMeta.getDisplayName(), cellValue);

                    //校验不成功则记录错误消息
                    if (result.getStatus() == false) {
                        String colName = cellName.replaceAll("\\d+", "");
                        ErrorMeta.ErrorMetaBuilder errorMetaBuilder = ErrorMeta.builder();
                        errorMetaBuilder.fieldName(columnMeta.getFieldName());
                        errorMetaBuilder.colName(colName);
                        errorMetaBuilder.rowName(cellName.replace(colName, ""));
                        errorMetaBuilder.message(result.getMessage());

                        context.getErrorMetas().add(errorMetaBuilder.build());
                    }
                } else {
                    throw new RuntimeException("校验器" + validatorBean + "未找到，请检查Bean配置！");
                }
            });
        }
    }


    private IExcelImportValidator getExcelImportValidator(IExcelContext context, String consumerBean) {
        ApplicationContext applicationContext = context.getApplicationContext();
        return applicationContext.getBean(consumerBean, IExcelImportValidator.class);
    }
}
