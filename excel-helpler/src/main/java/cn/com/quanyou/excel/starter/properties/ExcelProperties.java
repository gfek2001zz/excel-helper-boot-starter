package cn.com.quanyou.excel.starter.properties;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;

@Data
@ConfigurationProperties(prefix = "excel")
public class ExcelProperties {

    private String configType;
}
