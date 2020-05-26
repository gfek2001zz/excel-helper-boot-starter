package cn.com.quanyou.excel.entity;

public enum ExcelTypeEnum {

    XLS("xls"),
    XLSX("xlsx");

    private String value;

    ExcelTypeEnum(String value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return value;
    }
}
