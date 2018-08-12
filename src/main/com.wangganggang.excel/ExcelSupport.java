package com.wangganggang.excel;

import javax.servlet.http.HttpServletRequest;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月05日 下午11:01
 */
public class ExcelSupport {

    /**
     * 对文件流输出下载的中文文件名进行编码 屏蔽各种浏览器版本的差异性
     *
     * @param request
     * @param fileName
     * @return
     */
    public static String encodeChineseDownloadFileName(HttpServletRequest request, String fileName) {
        String agent = request.getHeader("USER-AGENT");
        try {
            if (null != agent && -1 != agent.indexOf("MSIE")) {
                fileName = URLEncoder.encode(fileName, "utf-8");
            } else {
                fileName = new String(fileName.getBytes("utf-8"), "iso8859-1");
            }
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return fileName;
    }

    /**
     * 判断对象是否为NotEmpty(!null或元素>0)<br>
     * 实用于对如下对象做判断:String Collection及其子类 Map及其子类
     *
     * @param object 待检查对象
     * @return boolean 返回的布尔值
     */
    public static boolean isNotEmpty(Object object) {
        return !isEmpty(object);
    }

    /**
     * 判断对象是否Empty(null或元素为0)<br>
     * 实用于对如下对象做判断:String Collection及其子类 Map及其子类
     *
     * @param object 待检查对象
     * @return boolean 返回的布尔值
     */
    public static boolean isEmpty(Object object) {
        if (object == null) {
            return true;
        }
        if (object == "") {
            return true;
        }
        if (object instanceof String) {
            if (((String) object).length() == 0) {
                return true;
            }
        } else if (object instanceof Collection) {
            if (((Collection) object).size() == 0) {
                return true;
            }
        } else if (object instanceof Map) {
            if (((Map) object).size() == 0) {
                return true;
            }
        } else if (object instanceof List[]) {
            for (List l : (List[]) object) {
                if (l.size() > 0) {
                    return true;
                }
            }
            return false;
        }
        return false;
    }
}
