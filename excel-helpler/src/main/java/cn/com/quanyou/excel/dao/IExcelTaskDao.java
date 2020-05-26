package cn.com.quanyou.excel.dao;

import cn.com.quanyou.excel.context.impl.ErrorMeta;
import cn.com.quanyou.excel.entity.*;
import org.apache.ibatis.annotations.Param;
import org.springframework.cache.annotation.CacheEvict;
import org.springframework.cache.annotation.Cacheable;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface IExcelTaskDao {
    void initTask(TaskModel data);

    void updateTaskStatus(@Param("status") TaskStatusEnum status, @Param("taskId") Long taskId);

    TaskModel findExcelTaskByTaskId(@Param("taskId") Long taskId);

    void updateTaskInfo(TaskModel taskModel);

    List<ErrorMeta> findErrorMeta(@Param("taskId") Long taskId);

    void deleteErrorMeta(@Param("taskId") Long taskId);

    @Cacheable(cacheNames = "excelCache", key = "'excelCfg'.concat(#p0).concat(#p1)", unless = "#result == null")
    ExcelCfgModel findExcelCfgByReportType(@Param("reportName") String reportName, @Param("reportType") String reportType);

    @Cacheable(cacheNames = "excelCache", key = "'excelSheetCfg'.concat(#p0).concat(#p1)", unless = "#result == null || #result.size() == 0")
    List<ExcelSheetCfgModel> findExcelSheetCfgByReportType(@Param("reportName") String reportName, @Param("reportType") String reportType);

    @Cacheable(cacheNames = "excelCache", key = "'excelColumnCfg'.concat(#p0).concat(#p1)", unless = "#result == null || #result.size() == 0")
    List<ExcelColumnCfgModel> findExcelColumnCfgByReportType(@Param("reportName") String reportName, @Param("reportType") String reportType);
}
