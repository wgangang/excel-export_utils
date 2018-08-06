package com.wangganggang.excel.utils;

import com.wangganggang.excel.ExcelConstants;
import com.wangganggang.excel.ExcelFiller;
import com.wangganggang.excel.ExcelSupport;
import com.wangganggang.excel.ExcelTemplate;
import jxl.Cell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import lombok.Getter;
import lombok.NoArgsConstructor;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月06日 上午12:42
 */
@NoArgsConstructor
public class ExcelFillerViaPOIUtils extends ExcelFiller {

    private WritableSheet wSheet = null;
    private HSSFSheet hSheet = null;
    /**
     * 创建excel文件
     */
    private WritableWorkbook wwb = null;
    private int sheetIndex = 0;
    private String sheetName = "";
    @Getter
    private HSSFWorkbook hwb = null;
    private ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

    /**
     * 用户自定义样式：false表示使用模板原有样式；true表示使用自定义的样式
     */
    @Getter
    private boolean customStyle = true;

    private HSSFCellStyle bodyHCs = null;
    private HSSFCellStyle titleHCs = null;
    private HSSFFont bodyHFont = null;
    private HSSFFont titleHFont = null;

    public ExcelFillerViaPOIUtils(ExcelTemplate pExcelTemplate) {
        setExcelTemplate(pExcelTemplate);
        BufferedInputStream bis = new BufferedInputStream(ExcelTemplate.class.getResourceAsStream(getExcelTemplate().getTemplatePath()));
        POIFSFileSystem fs = null;
        try {
            fs = new POIFSFileSystem(bis);
            hwb = new HSSFWorkbook(fs);
        } catch (IOException e) {
            logger.error("基于模板生成可写工作表出错了!");
            e.printStackTrace();
        }
    }

    public void fill(int i) {
        try {
            hSheet = hwb.getSheetAt(i);
            fillStatics(i);
            fillParameters(i);
            fillFields(i);
        } catch (Exception e) {
            logger.error("基于模板生成可写工作表出错了,错误信息:{}", e.getMessage());
            e.printStackTrace();
        }

    }

    public void write() {
        try {
            hwb.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public void fillStatics() {
        fillStatics(null);
    }

    public void fillStatics(Integer sheetIdx) {
        List statics = null;
        if (getExcelTemplate().isMultiParamTemplate()) {
            statics = getExcelTemplate().getStaticObject(sheetIdx);
        } else {
            statics = getExcelTemplate().getStaticObject();
        }
        for (int i = 0; i < statics.size(); i++) {
            Cell cell = (Cell) statics.get(i);
            HSSFRow hRow = hSheet.getRow(cell.getRow());
            if (hRow == null) {
                hRow = hSheet.createRow(cell.getRow());
            }
            HSSFCell hCell = hRow.getCell(cell.getColumn());
            if (hCell == null) {
                hCell = hRow.createCell(cell.getColumn());
            }
            hCell.setCellValue(cell.getContents());
            if (this.isCustomStyle()) {
                hCell.setCellStyle(getTitleHSSFCellStyle());
            }
        }
    }

    public void fillParameters() {
        fillParameters(null);
    }

    /**
     * 写入参数对象
     */
    public void fillParameters(Integer sheetIdx) {
        List parameters = null;
        if (getExcelTemplate().isMultiParamTemplate()) {
            parameters = getExcelTemplate().getParameterObject(sheetIdx);
        } else {
            parameters = getExcelTemplate().getParameterObject();
        }

        Map parameterMap = getExcelData().getParametersMap();
        for (int i = 0; i < parameters.size(); i++) {
            Cell cell = (Cell) parameters.get(i);
            String key = getKey(cell.getContents().trim());
            String type = getType(cell.getContents().trim());
            try {
                if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                    HSSFRow hRow = hSheet.getRow(cell.getRow());
                    if (hRow == null) {
                        hRow = hSheet.createRow(cell.getRow());
                    }
                    HSSFCell hCell = hRow.getCell(cell.getColumn());
                    if (hCell == null) {
                        hCell = hRow.createCell(cell.getColumn());
                    }
                    hCell.setCellValue(getDoubleValue(parameterMap.get(key)));
                    // number.setCellFormat(null);//getBodyCellStyle()
                    if (this.isCustomStyle()) {
                        hCell.setCellStyle(getBodyHSSFCellStyle());
                    }

                } else {
                    HSSFRow hRow = hSheet.getRow(cell.getRow());
                    if (hRow == null) {
                        hRow = hSheet.createRow(cell.getRow());
                    }
                    HSSFCell hCell = hRow.getCell(cell.getColumn());
                    if (hCell == null) {
                        hCell = hRow.createCell(cell.getColumn());
                    }
                    hCell.setCellValue(getString(parameterMap.get(key)));
                    // label.setCellFormat(null);
                    if (this.isCustomStyle()) {
                        hCell.setCellStyle(getBodyHSSFCellStyle());
                    }

                }
            } catch (Exception e) {
                logger.error("写入表格参数对象发生错误,错误信息:{}", e.getMessage());
                e.printStackTrace();
            }
        }
    }

    public void fillFields() throws Exception {
        fillFields(null);
    }

    public void fillFields(Integer sheetIdx) throws Exception {
        List fields = null;
        if (getExcelTemplate().isMultiParamTemplate()) {
            fields = getExcelTemplate().getFieldObject(sheetIdx);
        } else {
            fields = getExcelTemplate().getFieldObject();
        }
        List fieldList = getExcelData().getFieldsList();
        for (int j = 0; j < fieldList.size(); j++) {
            Map dataMap = new HashMap<>();
            HSSFRow hRow = null;
            Object object = fieldList.get(j);
            if (object instanceof Map<?, ?>) {
                Map domain = (Map) object;
                dataMap.putAll(domain);
            } else {
                logger.error("不支持的数据类型!");
            }
            for (int i = 0; i < fields.size(); i++) {
                Cell cell = (Cell) fields.get(i);
                String key = getKey(cell.getContents().trim());
                String type = getType(cell.getContents().trim());
                if (hRow == null) {
                    hRow = hSheet.createRow(cell.getRow() + j);
                }
                try {
                    if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                        BigDecimal val = BigDecimal.valueOf(getDoubleValue(dataMap.get(key)));
                        if (val == null) {
                            HSSFCell hCell = hRow.createCell(cell.getColumn());
                            hCell.setCellValue("");
                            if (this.isCustomStyle()) {
                                hCell.setCellStyle(getBodyHSSFCellStyle());
                            }
                        } else {
                            HSSFCell hCell = hRow.createCell(cell.getColumn());
                            hCell.setCellValue(val.doubleValue());
                            if (this.isCustomStyle()) {
                                hCell.setCellStyle(getBodyHSSFCellStyle());
                            }
                        }
                    } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_FORMULA)) {
                        HSSFCell hCell = hRow.createCell(cell.getColumn());
                        hCell.setCellFormula(key);
                        if (this.isCustomStyle()) {
                            hCell.setCellStyle(getBodyHSSFCellStyle());
                        }
                    } else {
                        HSSFCell hCell = hRow.createCell(cell.getColumn());
                        hCell.setCellValue(getString(dataMap.get(key)));
                        if (this.isCustomStyle()) {
                            hCell.setCellStyle(getBodyHSSFCellStyle());
                        }
                    }
                } catch (Exception e) {
                    logger.error("写入表格字段对象发生错误,错误信息:{}", e.getMessage());
                    e.printStackTrace();
                }
            }
        }
        int row = 0;
        row += fieldList.size();
        if (ExcelSupport.isEmpty(fieldList)) {
            if (ExcelSupport.isNotEmpty(fields)) {
                Cell cell = (Cell) fields.get(0);
                row = cell.getRow();
                wSheet.removeRow(row + 5);
                wSheet.removeRow(row + 4);
                wSheet.removeRow(row + 3);
                wSheet.removeRow(row + 2);
                wSheet.removeRow(row + 1);
                wSheet.removeRow(row);
            }
        } else {
            Cell cell = (Cell) fields.get(0);
            row += cell.getRow();
            fillVariables(wSheet, row);
        }
    }

    public void fillVariables(WritableSheet wSheet, int row) {
        List variables = getExcelTemplate().getVariableObject();
        if (variables != null) {
            Map parameterMap = getExcelData().getParametersMap();
            for (int i = 0; i < variables.size(); i++) {
                Cell cell = (Cell) variables.get(i);
                String key = getKey(cell.getContents().trim());
                String type = getType(cell.getContents().trim());
                try {
                    if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                        HSSFCell hCell = hSheet.getRow(cell.getRow()).getCell(cell.getColumn());
                        hCell.setCellValue(getDoubleValue(parameterMap.get(key)));
                        if (this.isCustomStyle()) {
                            hCell.setCellStyle(getTitleHSSFCellStyle());
                        }
                    } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_FORMULA)) {
                        /*
                         * 转义,$R表示当前行数
                         */
                        if (key.indexOf("$R") > -1) {
                            key = key.replaceAll("\\$R", String.valueOf(row));
                        }

                        HSSFCell hCell = hSheet.getRow(cell.getRow()).getCell(cell.getColumn());
                        hCell.setCellFormula(key);
                        if (this.isCustomStyle()) {
                            hCell.setCellStyle(getTitleHSSFCellStyle());
                        }
                    } else {
                        String content = getString(parameterMap.get(key));
                        if (ExcelSupport.isEmpty(content) && !key.equalsIgnoreCase("nbsp")) {
                            content = key;
                        }
                        HSSFCell hCell = hSheet.getRow(cell.getRow()).getCell(cell.getColumn());
                        hCell.setCellValue(content);
                        if (this.isCustomStyle()) {
                            hCell.setCellStyle(getTitleHSSFCellStyle());
                        }
                    }
                } catch (Exception e) {
                    logger.error("写入表格变量对象发生错误,错误信息:{}", e.getMessage());
                    e.printStackTrace();
                }
            }
        }
    }

    public HSSFCellStyle getBodyHSSFCellStyle() {
        if (this.bodyHCs == null) {
            this.bodyHCs = hwb.createCellStyle();
//			this.bodyHCs.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);
//			this.bodyHCs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            this.bodyHCs.setFont(getBodyHFont());
            this.bodyHCs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            this.bodyHCs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            this.bodyHCs.setBorderRight(HSSFCellStyle.BORDER_THIN);
            this.bodyHCs.setBorderTop(HSSFCellStyle.BORDER_THIN);
        }
        return this.bodyHCs;
    }

    public HSSFCellStyle getTitleHSSFCellStyle() {
        if (this.titleHCs == null) {
            this.titleHCs = hwb.createCellStyle();
//			titleHCs.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);
//			titleHCs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            this.titleHCs.setFont(getBodyHFont());
            this.titleHCs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            this.titleHCs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            this.titleHCs.setBorderRight(HSSFCellStyle.BORDER_THIN);
            this.titleHCs.setBorderTop(HSSFCellStyle.BORDER_THIN);
        }
        return this.titleHCs;
    }

    public void addSheet(String sheetName) {
        wSheet = wwb.createSheet(sheetName, sheetIndex++);
    }

    public HSSFFont getBodyHFont() {
        if (this.bodyHFont == null) {
            this.bodyHFont = hwb.createFont();
            this.bodyHFont.setFontName("宋体");
        }
        return this.bodyHFont;
    }
}
