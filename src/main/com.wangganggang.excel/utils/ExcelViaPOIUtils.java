package com.wangganggang.excel.utils;

import com.wangganggang.excel.ExcelData;
import com.wangganggang.excel.ExcelExporter;
import com.wangganggang.excel.ExcelSupport;
import com.wangganggang.excel.ExcelTemplate;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月12日 下午5:36
 */
@NoArgsConstructor
@Getter
@Setter
public class ExcelViaPOIUtils extends ExcelExporter {

    private String filename = "Excel.xls";

    private String sheetName = "";

    private List list = new ArrayList();

    /**
     * 用户自定义样式：false表示默认；true表示使用自定义的样式
     */
    private Boolean[] customStyle;

    /**
     * 是否使用多个sheet来定义样本
     */
    private boolean multiParamTemplate = false;

    /**
     * 是否样本sheet个数
     */
    private int paramSheetNums = 1;

    public ExcelViaPOIUtils(List<Map> list) {
        this.list = list;
    }

    /**
     * 导出Excel
     *
     * @return
     * @throws IOException
     */
    public HSSFWorkbook export() throws IOException {

        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setTemplatePath(getTemplatePath());
        excelTemplate.setMultiParamTemplate(multiParamTemplate);
        excelTemplate.setParamSheetNum(paramSheetNums);
        excelTemplate.parse();

        ExcelFillerViaPOIUtils excelFillerViaPOIUtils = new ExcelFillerViaPOIUtils(excelTemplate);
        for (int i = 0; i < this.list.size(); i++) {
            Map map = (Map) list.get(i);
            Map dto = (Map) map.get("parametersMap");
            List fList = (List) map.get("fieldsList");
            ExcelData excelData = new ExcelData(dto, fList);
            excelFillerViaPOIUtils.setExcelData(excelData);
            if (getSheetName() != null && !"".equals(getSheetName())) {
                excelFillerViaPOIUtils.setSheetName(getSheetName());
            }
            if (getCustomStyle(i) != null) {
                excelFillerViaPOIUtils.setCustomStyle(getCustomStyle(i));
            }
            excelFillerViaPOIUtils.fill(i);
        }
        return excelFillerViaPOIUtils.getHwb();
    }

    /**
     * 导出Excel
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @Override
    public void export(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        filename = ExcelSupport.encodeChineseDownloadFileName(request, getFilename());
        response.setHeader("Content-Disposition", "attachment; filename=" + filename + ";");

        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setTemplatePath(getTemplatePath());
        excelTemplate.setMultiParamTemplate(multiParamTemplate);
        excelTemplate.setParamSheetNum(paramSheetNums);
        excelTemplate.parse();

        // ExcelFiller excelFiller = new ExcelFiller(excelTemplate, excelData);
        ExcelFillerViaPOIUtils excelFillerViaPOIUtils = new ExcelFillerViaPOIUtils(excelTemplate);
        ByteArrayOutputStream bos = null;
        ServletOutputStream os = null;
        for (int i = 0; i < this.list.size(); i++) {
            Map map = (Map) list.get(i);
            Map dto = (Map) map.get("parametersMap");
            List fList = (List) map.get("fieldsList");
            ExcelData excelData = new ExcelData(dto, fList);
            excelFillerViaPOIUtils.setExcelData(excelData);
            if (getSheetName() != null && !"".equals(getSheetName())) {
                excelFillerViaPOIUtils.setSheetName(getSheetName());
            }
            if (getCustomStyle(i) != null) {
                excelFillerViaPOIUtils.setCustomStyle(getCustomStyle(i));
            }
            excelFillerViaPOIUtils.fill(i);
        }
        excelFillerViaPOIUtils.write();
        bos = excelFillerViaPOIUtils.getOutputStream();
        os = response.getOutputStream();
        os.write(bos.toByteArray());
        os.flush();
        os.close();
    }

    /**
     * 导出Excel
     *
     * @param request
     * @param response
     * @throws IOException
     */
    public void export(HttpServletRequest request, HttpServletResponse response, ExcelTemplate excelTemplate) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        filename = ExcelSupport.encodeChineseDownloadFileName(request, getFilename());
        response.setHeader("Content-Disposition", "attachment; filename=" + filename + ";");

        excelTemplate.setTemplatePath(getTemplatePath());
        excelTemplate.setMultiParamTemplate(multiParamTemplate);
        excelTemplate.setParamSheetNum(paramSheetNums);
        excelTemplate.parse();

        // ExcelFiller excelFiller = new ExcelFiller(excelTemplate, excelData);
        ExcelFillerUtils excelFillerUtils = new ExcelFillerUtils(excelTemplate);
        ByteArrayOutputStream bos = null;
        ServletOutputStream os = null;
        for (int i = 0; i < this.list.size(); i++) {
            Map map = (Map) list.get(i);
            Map dto = (Map) map.get("parametersMap");
            List fList = (List) map.get("fieldsList");
            ExcelData excelData = new ExcelData(dto, fList);
            excelFillerUtils.setExcelData(excelData);
            // ByteArrayOutputStream bos = excelFiller.fill();
            if (getSheetName() != null && !"".equals(getSheetName())) {
                excelFillerUtils.setSheetName(getSheetName());
            }
            if (getCustomStyle(i) != null) {
                excelFillerUtils.setCustomStyle(getCustomStyle(i));
            }

            excelFillerUtils.fill(i + excelTemplate.getSkipSheets());
        }
        //删除多余的sheet row column
        if (excelTemplate.getCleanTheDoc() != null) {
            excelFillerUtils.cleanTheDoc(excelTemplate.getCleanTheDoc());
        }
        excelFillerUtils.write();
        bos = excelFillerUtils.getOutputStream();
        os = response.getOutputStream();
        os.write(bos.toByteArray());
        os.flush();
        os.close();
    }

    public Boolean getCustomStyle(int i) {
        return this.customStyle == null || this.customStyle.length == 0 ? null : this.customStyle[i];
    }
}
