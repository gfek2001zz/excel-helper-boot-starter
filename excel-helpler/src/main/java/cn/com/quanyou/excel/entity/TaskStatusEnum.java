package cn.com.quanyou.excel.entity;

public enum TaskStatusEnum {

    PENDING("PENDING"),
    PROCESSING("PROCESSING"),
    PROCESSED("PROCESSED"),
    FAILED("FAILED"),
    ERROR("ERROR");

    private String value;
    TaskStatusEnum(String value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return this.value;
    }
}
