package cn.com.quanyou.excel.entity;

import lombok.Data;

import java.util.List;

@Data
public class Page<T> {
    private int total;
    private List<T> records;
}
