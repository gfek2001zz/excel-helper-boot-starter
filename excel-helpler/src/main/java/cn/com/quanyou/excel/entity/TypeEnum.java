package cn.com.quanyou.excel.entity;

public enum TypeEnum {
    IMPORT("IMPORT", "configs/%s.import.xml"),
    EXPORT("EXPORT", "configs/%s.export.xml");

    private String value;
    private String template;
    TypeEnum(String value, String template) {
        this.value = value;
        this.template = template;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getTemplate() {
        return template;
    }

    public void setTemplate(String template) {
        this.template = template;
    }
}
