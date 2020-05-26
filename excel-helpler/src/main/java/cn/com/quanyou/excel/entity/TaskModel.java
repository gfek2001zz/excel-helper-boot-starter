package cn.com.quanyou.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class TaskModel {

    private Long taskId;
    private String reportType;
    private String fileName;
    private String downloadUrl;
    private String operationType;
    private String execResults;
    private TaskStatusEnum status;
    private Date startTime;
    private Date endTime;

    private String createUser;
    private Date createTime;
    private String lastUpdateUser;
    private Date lastUpdateTime;
}
