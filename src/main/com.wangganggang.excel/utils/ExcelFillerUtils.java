package com.wangganggang.excel.utils;

import com.wangganggang.excel.ExcelConstants;
import com.wangganggang.excel.ExcelFiller;
import com.wangganggang.excel.ExcelSupport;
import com.wangganggang.excel.ExcelTemplate;
import jxl.Cell;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月06日 上午12:11
 */
@NoArgsConstructor
public class ExcelFillerUtils extends ExcelFiller {

    /**
     * 创建excel文件
     */
    private WritableWorkbook wwb = null;

    /**
     * 工作表
     */
    private WritableSheet wSheet = null;
    private int sheetIndex = 0;

    private Workbook wb;

    @Getter
    private ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

    /**
     * 用户自定义样式：false表示使用模板原有样式；true表示使用自定义的样式
     */
    @Setter
    @Getter
    private boolean customStyle = true;

    @Setter
    @Getter
    private String sheetName = "sheet";

    public ExcelFillerUtils(ExcelTemplate pExcelTemplate) {
        setExcelTemplate(pExcelTemplate);
        InputStream is = ExcelTemplate.class.getResourceAsStream(getExcelTemplate().getTemplatePath());
        try {
            wb = Workbook.getWorkbook(is);
            WorkbookSettings settings = new WorkbookSettings();
            settings.setWriteAccess(null);
            wwb = Workbook.createWorkbook(outputStream, wb, settings);
        } catch (BiffException e) {
            logger.error("基于模板生成可写工作表出错了,错误信息:{}", e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            logger.error("基于模板生成可写工作表出错了,错误信息:{}", e.getMessage());
            e.printStackTrace();
        }
    }

    public void fill(int i) {
        try {
            wSheet = wwb.getSheet(i);
            if (!"".equals(this.sheetName) && this.sheetName != null) {
                wSheet.setName(this.sheetName);
            }
            // this.addSheet(getSheetName());
            fillStatics(wSheet, i);
            fillParameters(wSheet, i);
            fillFields(wSheet, i);
        } catch (Exception e) {
            logger.error("基于模板生成可写工作表出错了,错误信息:{}", e.getMessage());
            e.printStackTrace();
        }
    }



    public void write() {
        try {
            wwb.write();
            wwb.close();
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }

    }

    public void fillStatics(WritableSheet wSheet) {
        fillStatics(wSheet, null);
    }

    public void fillStatics(WritableSheet wSheet, Integer sheetIdx) {
        List statics = null;
        if (getExcelTemplate().isMultiParamTemplate()) {
            statics = getExcelTemplate().getStaticObject(sheetIdx);
        } else {
            statics = getExcelTemplate().getStaticObject();
        }
        for (int i = 0; i < statics.size(); i++) {
            Cell cell = (Cell) statics.get(i);
            Label label = new Label(cell.getColumn(), cell.getRow(), cell.getContents());
            Cell thisCell = wSheet.getCell(cell.getColumn(), cell.getRow());
            if (this.isCustomStyle()) {
                label.setCellFormat(getTitleCellStyle());
            } else {
                label.setCellFormat(thisCell.getCellFormat());
            }

            try {
                wSheet.addCell(label);
            } catch (Exception e) {
                logger.error("写入静态对象发生错误,错误信息:{}", e.getMessage());
                e.printStackTrace();
            }
        }
    }

    public void fillParameters(WritableSheet wSheet) {
        fillParameters(wSheet, null);
    }

    public void fillParameters(WritableSheet wSheet, Integer sheetIdx) {
        List parameters = null;
        if (getExcelTemplate().isMultiParamTemplate()) {
            parameters = getExcelTemplate().getParameterObjct(sheetIdx);
        } else {
            parameters = getExcelTemplate().getParameterObject();
        }
        Map parameterMap = getExcelData().getParametersMap();
        for (int i = 0; i < parameters.size(); i++) {
            Cell cell = (Cell) parameters.get(i);
            String key = getKey(cell.getContents().trim());
            String type = getType(cell.getContents().trim());
            Cell thisCell = wSheet.getCell(cell.getColumn(), cell.getRow());
            try {
                if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                    Number number = new Number(cell.getColumn(), cell.getRow(), getDoubleValue(parameterMap.get(key)));
                    // number.setCellFormat(null);//getBodyCellStyle()
                    if (this.isCustomStyle()) {
                        number.setCellFormat(getBodyCellStyle());
                    } else {
                        number.setCellFormat(thisCell.getCellFormat());
                    }
                    wSheet.addCell(number);
                } else {
                    Label label = new Label(cell.getColumn(), cell.getRow(), getString(parameterMap.get(key)));
                    // label.setCellFormat(null);
                    if (this.isCustomStyle()) {
                        label.setCellFormat(getBodyCellStyle());
                    } else {
                        label.setCellFormat(thisCell.getCellFormat());
                    }
                    wSheet.addCell(label);
                }
            } catch (Exception e) {
                logger.error("写入表格参数对象发生错误,错误信息:{}", e.getMessage());
                e.printStackTrace();
            }
        }
    }

    public void fillFields(WritableSheet wSheet) throws Exception {
        fillFields(wSheet, null);
    }

    public void fillFields(WritableSheet wSheet, Integer sheetIdx) throws Exception {
        List fields = null;
        if (getExcelTemplate().isMultiParamTemplate()) {
            fields = getExcelTemplate().getFieldObject(sheetIdx);
        } else {
            fields = getExcelTemplate().getFieldObject();
        }
        List fieldList = getExcelData().getFieldsList();
        for (int j = 0; j < fieldList.size(); j++) {
            Map dataMap = new HashMap<>();
            Object object = fieldList.get(j);
            if (object instanceof Map<?, ?>) {
                Map dto = (Map) object;
                dataMap.putAll(dto);
            } else {
                logger.error("不支持的数据类型!");
            }
            for (int i = 0; i < fields.size(); i++) {
                Cell cell = (Cell) fields.get(i);
                String key = getKey(cell.getContents().trim());
                String type = getType(cell.getContents().trim());
                Cell thisCell = wSheet.getCell(cell.getColumn(), cell.getRow() + j);
                try {

                    if ("Percent".equals(getString(dataMap.get("RenderType")))) {
                        Label label = new Label(cell.getColumn(), cell.getRow() + j, getString(dataMap.get(key)));
                        // label.setCellFormat(null);
                        if (this.isCustomStyle()) {
                            label.setCellFormat(getBodyCellStyle());
                        } else {
                            label.setCellFormat(thisCell.getCellFormat());
                        }

                        wSheet.addCell(label);
                    } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                        BigDecimal val = BigDecimal.valueOf(getDoubleValue(dataMap.get(key)));
                        if (val == null) {
                            Label label = new Label(cell.getColumn(), cell.getRow() + j, "");
                            // label.setCellFormat(null);
                            if (this.isCustomStyle()) {
                                label.setCellFormat(getBodyCellStyle());
                            } else {
                                label.setCellFormat(thisCell.getCellFormat());
                            }

                            wSheet.addCell(label);
                        } else {
                            Number number = new Number(cell.getColumn(), cell.getRow() + j, val.doubleValue());
                            // number.setCellFormat(null);
                            if (this.isCustomStyle()) {
                                number.setCellFormat(getBodyCellStyle());
                            } else {
                                number.setCellFormat(thisCell.getCellFormat());
                            }
                            wSheet.addCell(number);
                        }
                    } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_FORMULA)) {
                        Formula formula = new Formula(cell.getColumn(), cell.getRow() + j, key);
                        wSheet.addCell(formula);
                    } else {
                        Label label = new Label(cell.getColumn(), cell.getRow() + j, getString(dataMap.get(key)));
                        // label.setCellFormat(null);
                        if (this.isCustomStyle()) {
                            label.setCellFormat(getBodyCellStyle());
                        } else {
                            label.setCellFormat(thisCell.getCellFormat());
                        }

                        wSheet.addCell(label);
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
                Cell thisCell = wSheet.getCell(cell.getColumn(), row);
                try {
                    if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                        Number number = new Number(cell.getColumn(), row, getDoubleValue(parameterMap.get(key)));
                        number.setCellFormat(getTitleCellStyle());
                        wSheet.addCell(number);
                    } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_FORMULA)) {
                        /*
                         * 转义,$R表示当前行数
                         */
                        if (key.indexOf("$R") > -1) {
                            key = key.replaceAll("\\$R", String.valueOf(row));
                        }
                        Formula formula = new Formula(cell.getColumn(), row, key);
                        // Formula formula = new Formula(cell.getColumn(),
                        // row,"SUM(E4:E4)" );
                        // formula.setCellFormat(null);
                        if (this.isCustomStyle()) {
                            formula.setCellFormat(getTitleCellStyle());
                        } else {
                            formula.setCellFormat(thisCell.getCellFormat());
                        }
                        wSheet.addCell(formula);
                    } else {
                        String content = getString(parameterMap.get(key));
                        if (ExcelSupport.isEmpty(content) && !key.equalsIgnoreCase("nbsp")) {
                            content = key;
                        }
                        Label label = new Label(cell.getColumn(), row, content);
                        if (this.isCustomStyle()) {
                            label.setCellFormat(getTitleCellStyle());
                        } else {
                            label.setCellFormat(thisCell.getCellFormat());
                        }
                        wSheet.addCell(label);
                    }
                } catch (Exception e) {
                    logger.error("写入表格变量对象发生错误,错误信息:{}", e.getMessage());
                    e.printStackTrace();
                }
            }
        }
    }

    @Override
    public WritableCellFormat getBodyCellStyle() {
        /**
         * WritableFont.createFont("宋体")：设置字体为宋体
         * 10：设置字体大小 WritableFont.NO_BOLD:设置字体非加粗
         * （BOLD：加粗 NO_BOLD：不加粗） false：设置非斜体
         * UnderlineStyle.NO_UNDERLINE：没有下划线
         */
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"), 10, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE);
        WritableCellFormat bodyFormat = new WritableCellFormat(font);
        try {
            // 设置单元格背景色：表体为白色
            // bodyFormat.setBackground(Colour.WHITE);
            // 设置表头表格边框样式
            // 整个表格线为细线、黑色
            bodyFormat.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
            bodyFormat.setWrap(true);
        } catch (WriteException e) {
            System.out.println("表体单元格样式设置失败！");
        }
        return bodyFormat;
    }

    @Override
    public WritableCellFormat getTitleCellStyle() {
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"), 10, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE);
        WritableCellFormat bodyFormat = new WritableCellFormat(font);
        try {
            // 设置单元格背景色：表体为白色
            // bodyFormat.setBackground(Colour.WHITE);
            // 设置表头表格边框样式
            // 整个表格线为细线、黑色
            bodyFormat.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
            bodyFormat.setWrap(true);
            // //设置垂直对齐为居中对齐
            bodyFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            // 水平对齐
            bodyFormat.setAlignment(jxl.format.Alignment.CENTRE);
        } catch (WriteException e) {
            System.out.println("表体单元格样式设置失败！");
        }
        return bodyFormat;
    }

    public void addSheet(String sheetName) {
        wSheet = wwb.createSheet(sheetName, sheetIndex++);
    }
}
