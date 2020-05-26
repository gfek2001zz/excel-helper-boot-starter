package cn.com.quanyou.excel.context.utils;

import cn.com.quanyou.excel.context.IExcelContext;
import lombok.extern.slf4j.Slf4j;
import org.xml.sax.Attributes;

import java.io.File;
import java.lang.reflect.Method;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.beanutils.BeanUtils;

@Slf4j
public class EntityUtils {

    public static <T> T getProperty(Object obj, String name) {
        T result = null;

        try {
            StringBuffer methodName = new StringBuffer();
            String prefix = name.substring(0, 1).toUpperCase();

            methodName.append("get").append(prefix).append(name.substring(1));
            Method method = obj.getClass().getDeclaredMethod(methodName.toString(), new Class[]{});

            result = (T) method.invoke(obj, null);
        } catch (Exception ex) {
            log.error(ex.getMessage(), ex);
        } finally {
            return result;
        }
    }


    public static String getExcelFilePath(IExcelContext context, String excelName) {
        String fileName = context.getFileName();
        Pattern pattern = Pattern.compile("\\#\\{(.*)\\}");
        Matcher matcher = pattern.matcher(excelName);
        if (matcher.find() && null != context.getParamMap()) {
            Map<String, Object> param = context.getParamMap();
            String fName = param.get(matcher.group(1)).toString();
            String source = excelName;
            String convert = source.replaceAll("\\#\\{(.*)\\}", fName);
            fileName =  fileName + File.separator + convert;
        } else {
            fileName =  fileName + File.separator + excelName;
        }

        return fileName;
    }


    public static void parseXML2Entity(Object entityObj, Attributes attributes) {
        int attrNum = attributes.getLength();
        for (int i = 0; i < attrNum; i ++) {
            String attrName = attributes.getQName(i);

            try {
                BeanUtils.setProperty(entityObj, attrName, attributes.getValue(i));
            } catch (Exception ex) {
                log.error(ex.getMessage(), ex);
            }
        }

        log.debug("end");
    }
}
