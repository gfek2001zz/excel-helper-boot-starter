package cn.com.quanyou.excel.exported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.exception.ExcelException;
import com.github.tobato.fastdfs.domain.fdfs.StorePath;
import com.github.tobato.fastdfs.service.FastFileStorageClient;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Slf4j
@Component
public class ExcelFileUploadHandler {

    @Autowired
    private FastFileStorageClient fastFileClient;

    @Value("${fdfs.downurl.prefix}")
    private String fdfsUrl;

    /**
     * 上传Excel文件
     *
     * @param wb
     * @return
     */
    public String uploadExcelFile(IExcelContext context, Workbook wb) {
        ByteArrayOutputStream byteOs = new ByteArrayOutputStream();

        try {
            wb.write(byteOs);
            return uploadExcelFile(context, byteOs);
        } catch (Exception ex) {
            throw new ExcelException(ex);
        } finally {
            if (byteOs != null) {
                try {
                    byteOs.close();
                } catch (IOException e) {
                }
            }
        }
    }

    public String uploadExcelFile(IExcelContext context, ByteArrayOutputStream byteOs) {
        ByteArrayInputStream byteIs = null;

        try {
            byteIs = new ByteArrayInputStream(byteOs.toByteArray());
            StorePath storePath = fastFileClient.uploadFile(byteIs, byteIs.available(), context.getExcelType(), null);return fdfsUrl + "/" + storePath.getFullPath();
        } catch (Exception ex) {
            throw new ExcelException(ex);
        } finally {
            if (byteOs != null) {
                try {
                    byteOs.close();
                } catch (IOException e) {
                }
            }
            if (byteIs != null) {
                try {
                    byteIs.close();
                } catch (IOException e) {
                }
            }
        }
    }
}
