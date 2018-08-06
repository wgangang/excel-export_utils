package com.wangganggang.excel;

import jxl.Cell;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.*;
import jxl.write.Number;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author wangganggang
 * @date 2018年08月05日 下午11:30
 * @desc Excel数据填充器
 */
@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class ExcelFiller {
    protected Logger logger = LogManager.getLogger(getClass());

    private ExcelTemplate excelTemplate = null;

    private ExcelData excelData = null;

    public ByteArrayOutputStream fill() {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            InputStream is = ExcelFiller.class.getResourceAsStream(getExcelTemplate().getTemplatePath());
            Workbook wb = Workbook.getWorkbook(is);
            WorkbookSettings settings = new WorkbookSettings();
            settings.setWriteAccess(null);
            WritableWorkbook wwb = Workbook.createWorkbook(bos, wb, settings);
            WritableSheet wSheet = wwb.getSheet(0);
            fillStatics(wSheet);
            fillParameters(wSheet);
            fillFields(wSheet);
            wwb.write();
            wwb.close();
            wb.close();
        } catch (Exception e) {
            logger.error("基于模板生成可写工作表出错了,错误信息:{}", e.getMessage());
            e.printStackTrace();
        }
        return bos;
    }

    /**
     * 写入静态对象
     *
     * @param wSheet
     */
    private void fillStatics(WritableSheet wSheet) {
        List statics = getExcelTemplate().getStaticObject();
        for (int i = 0; i < statics.size(); i++) {
            Cell cell = (Cell) statics.get(i);
            Label label = new Label(cell.getColumn(), cell.getRow(), cell.getContents());
            label.setCellFormat(cell.getCellFormat());
            try {
                wSheet.addCell(label);
            } catch (Exception e) {
                logger.error("写入静态对象发生错误,错误信息:{}", e.getMessage());
                e.printStackTrace();
            }
        }
    }

    /**
     * 写如参数对象
     *
     * @param wSheet
     */
    private void fillParameters(WritableSheet wSheet) {
        List parameters = getExcelTemplate().getParameterObject();
        Map parameterMap = getExcelData().getParametersMap();
        for (int i = 0; i < parameters.size(); i++) {
            Cell cell = (Cell) parameters.get(i);
            String key = getKey(cell.getContents().trim());
            String type = getType(cell.getContents().trim());
            try {
                if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                    Number number = new Number(cell.getColumn(), cell.getRow(), getDoubleValue(parameterMap.get(key)));
                    number.setCellFormat(getBodyCellStyle());
                    wSheet.addCell(number);
                } else {
                    Label label = new Label(cell.getColumn(), cell.getRow(), getString(parameterMap.get(key)));
                    label.setCellFormat(getBodyCellStyle());
                    wSheet.addCell(label);
                }
            } catch (Exception e) {
                logger.error("写入表格参数对象发生错误,错误信息:{}", e.getMessage());
                e.printStackTrace();
            }
        }
    }

    /**
     * 写入字段对象
     *
     * @param wSheet
     * @throws Exception
     */
    private void fillFields(WritableSheet wSheet) {
        List fields = getExcelTemplate().getFieldObject();
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
                try {
                    if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                        Number number = new Number(cell.getColumn(), cell.getRow() + j, getDoubleValue(dataMap.get(key)));
                        number.setCellFormat(getBodyCellStyle());
                        wSheet.addCell(number);
                    } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_FORMULA)) {
                        Formula formula = new Formula(cell.getColumn(), cell.getRow() + j, key);
                        wSheet.addCell(formula);
                    } else {
                        Label label = new Label(cell.getColumn(), cell.getRow() + j, getString(dataMap.get(key)));
                        label.setCellFormat(getBodyCellStyle());
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

    private void fillVariables(WritableSheet wSheet, int row) {
        List variables = getExcelTemplate().getVariableObject();
        Map parameterMap = getExcelData().getParametersMap();
        for (int i = 0; i < variables.size(); i++) {
            Cell cell = (Cell) variables.get(i);
            String key = getKey(cell.getContents().trim());
            String type = getType(cell.getContents().trim());
            try {
                if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_NUMBER)) {
                    Number number = new Number(cell.getColumn(), row, getDoubleValue(parameterMap.get(key)));
                    number.setCellFormat(getBodyCellStyle());
                    wSheet.addCell(number);
                } else if (type.equalsIgnoreCase(ExcelConstants.EXCELTPL_DATATYPE_FORMULA)) {
                    /*
                     * 转义,$R表示当前行数
                     */
                    if (key.indexOf("$R") > -1) {
                        key = key.replaceAll("\\$R", String.valueOf(row));
                    }
                    Formula formula = new Formula(cell.getColumn(), row, key);
                    // Formula formual = new Formula(cell.getColumn(),
                    // row,"SUM(E4:E4)" );
                    formula.setCellFormat(getBodyCellStyle());
                    wSheet.addCell(formula);
                } else {
                    String content = getString(parameterMap.get(key));
                    if (ExcelSupport.isEmpty(content) && !key.equalsIgnoreCase("nbsp")) {
                        content = key;
                    }
                    Label label = new Label(cell.getColumn(), row, content);
                    label.setCellFormat(getBodyCellStyle());
                    wSheet.addCell(label);
                }
            } catch (Exception e) {
                logger.error("写入表格变量对象发生错误,错误信息:{}", e.getMessage());
                e.printStackTrace();
            }
        }
    }

    /**
     * 删除多余的sheet row column(反序删，序号从大到小）
     *
     * @param cleanTheDoc sheetIndex:R|C|S:endRowIndex|endColumnIndex|endSheetIndex:deleteLength[,REPEAT]
     */
    public void cleanTheDoc(String cleanTheDoc, WritableWorkbook wwb) {
        if (cleanTheDoc == null) {
            return;
        }
        String[] commands = cleanTheDoc.split(",");
        if (commands.length == 0) {
            return;
        }
        for (int i = 0; i < commands.length; i++) {
            String[] command = commands[i].split(":", 4);
            String type = command[1];
            switch (type.charAt(0)) {
                case 'R':
                    WritableSheet sheetR = wwb.getSheet(Integer.valueOf(command[0]));
                    int endR = Integer.valueOf(command[2]);
                    int lengthR = Integer.valueOf(command[3]);
                    int startR = endR - lengthR;
                    for (int j = endR; j > startR; j--) {
                        sheetR.removeRow(j);
                    }
                    break;
                case 'C':
                    WritableSheet sheetC = wwb.getSheet(Integer.valueOf(command[0]));
                    int endC = Integer.valueOf(command[2]);
                    int lengthC = Integer.valueOf(command[3]);
                    int startC = endC - lengthC;
                    for (int j = endC; j > startC; j--) {
                        sheetC.removeColumn(j);
                    }
                    break;
                case 'S':
                    int endS = Integer.valueOf(command[2]);
                    int lengthS = Integer.valueOf(command[3]);
                    int startS = endS - lengthS;
                    for (int j = endS; j > startS; j--) {
                        wwb.removeSheet(j);
                    }
                    break;
                default:
                    break;
            }
        }
    }

    protected String getString(Object object) {
        if (Objects.isNull(object)) {
            return "";
        }
        return object.toString();
    }

    protected double getDoubleValue(Object object) {
        if (Objects.isNull(object)) {
            return 0D;
        }
        Double dobuleObject = new Double(object.toString());
        return dobuleObject.doubleValue();
    }

    /**
     * 单元格样式的设定
     *
     * @return
     */
    protected WritableCellFormat getBodyCellStyle() {
        /*
         * WritableFont.createFont("宋体")：设置字体为宋体 10：设置字体大小
         * WritableFont.NO_BOLD:设置字体非加粗（BOLD：加粗 NO_BOLD：不加粗） false：设置非斜体
         * UnderlineStyle.NO_UNDERLINE：没有下划线
         */
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"), 10,
                WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE);
        WritableCellFormat bodyFormat = new WritableCellFormat(font);
        try {
            // 设置单元格背景色：表体为白色
            bodyFormat.setBackground(jxl.format.Colour.WHITE);
            // 设置表头表格边框样式
            // 整个表格线为细线、黑色
            bodyFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN, Colour.BLACK);
            bodyFormat.setWrap(true);
        } catch (WriteException e) {
            System.out.println("表体单元格样式设置失败！");
        }
        return bodyFormat;
    }

    /**
     * 表头单元格样式的设定
     * 默认认为,如果是非内容部分,则为标题部分
     *
     * @return
     */
    protected WritableCellFormat getTitleCellStyle() {
        /*
         * WritableFont.createFont("宋体")：设置字体为宋体 10：设置字体大小
         * WritableFont.NO_BOLD:设置字体非加粗（BOLD：加粗 NO_BOLD：不加粗） false：设置非斜体
         * UnderlineStyle.NO_UNDERLINE：没有下划线
         */
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"), 12,
                WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE);
        WritableCellFormat bodyFormat = new WritableCellFormat(font);
        try {
            // 设置单元格背景色：表体为白色
            bodyFormat.setBackground(Colour.WHITE);
            // 设置表头表格边框样式
            // 整个表格线为细线、黑色
            bodyFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN, Colour.BLACK);
            bodyFormat.setWrap(true);
            ////设置垂直对齐为居中对齐
            bodyFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
            //水平对齐
            bodyFormat.setAlignment(jxl.write.Alignment.CENTRE);
        } catch (WriteException e) {
            System.out.println("表体单元格样式设置失败！");
        }
        return bodyFormat;
    }

    /**
     * 获取模板单元格标记数据类型
     *
     * @param pType 模板元标记
     * @return 数据类型
     */
    protected static String getType(String pType) {
        String type = ExcelConstants.EXCELTPL_DATATYPE_LABEL;
        if (pType.indexOf(":n") != -1 || pType.indexOf(":N") != -1) {
            type = ExcelConstants.EXCELTPL_DATATYPE_NUMBER;
        }
        if (pType.indexOf(":u") != -1 || pType.indexOf(":U") != -1) {
            type = ExcelConstants.EXCELTPL_DATATYPE_FORMULA;
        }
        return type;
    }

    /**
     * 获取模板键名
     *
     * @param pKey
     * @return
     */
    protected static String getKey(String pKey) {
        String key = null;
        try {
            int index = pKey.lastIndexOf(":");
            if (index == -1) {
                key = pKey.substring(3, pKey.length() - 1);
            } else {
                key = pKey.substring(3, index);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return key;
    }
}
