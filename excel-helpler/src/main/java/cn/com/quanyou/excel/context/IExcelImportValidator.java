package cn.com.quanyou.excel.context;

import cn.com.quanyou.excel.entity.ValidatorResult;

public interface IExcelImportValidator {
     ValidatorResult validate(String fieldName, Object value);
}
