package cn.com.quanyou.excel.entity;

public enum ConfigTypeEnum {
    DATABASE("DATABASE", "数据库配置"),
    XML("XML", "XML配置");

    private String value;
    private String text;

    ConfigTypeEnum(String value, String text) {
        this.value = value;
        this.text = text;
    }

    @Override
    public String toString() {
        return this.value;
    }
}
