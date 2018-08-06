package com.wangganggang.excel;

import lombok.Getter;
import lombok.Setter;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月06日 上午12:00
 */
@Getter
@Setter
public class ExcelExporter {

    private String templatePath;
    private Map parametersMap;
    private List fieldsList;
    private String fileName = "Excel.xls";

    public void setData(Map pMap, List pList) {
        parametersMap = pMap;
        fieldsList = pList;
    }

    public void export(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        fileName = ExcelSupport.encodeChineseDownloadFileName(request, getFileName());
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ";");
        ExcelData excelData = new ExcelData(parametersMap, fieldsList);
        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setTemplatePath(getTemplatePath());
        excelTemplate.parse();
        ExcelFiller excelFiller = new ExcelFiller(excelTemplate, excelData);
        ByteArrayOutputStream bos = excelFiller.fill();
        ServletOutputStream os = response.getOutputStream();
        os.write(bos.toByteArray());
        os.flush();
        os.close();
    }

}
