package com.wangganggang.excel;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author wangganggang
 * @date 2018年08月05日 下午10:46
 */
@NoArgsConstructor
@Getter
@Setter
public class ExcelTemplate {
    protected Logger logger = LoggerFactory.getLogger(ExcelTemplate.class);

    /**
     *
     */
    private List staticObject = null;

    /**
     *
     */
    private List parameterObject = null;

    /**
     *
     */
    private List fieldObject = null;

    /**
     *
     */
    private List variableObject = null;

    /**
     *
     */
    private String templatePath = null;

    /**
     *
     */
    private boolean multiParamTemplate = false;

    /**
     *
     */
    private int paramSheetNum = 1;

    /**
     *
     */
    private int skipSheets = 0;

    /**
     *
     */
    private String cleanTheDoc = null;

    /**
     *
     */
    private List<HashMap<String, List>> mObjectList = new ArrayList<HashMap<String, List>>();

    public ExcelTemplate(String pTemplatePath) {
        templatePath = pTemplatePath;
    }

    public void parse() {
        if (ExcelSupport.isEmpty(templatePath)) {
            logger.error("Excel模板路径不能为空!");
            return;
        }

        InputStream is = ExcelTemplate.class.getResourceAsStream(templatePath);
        Workbook workbook = null;
        try {
            workbook = Workbook.getWorkbook(is);
        } catch (Exception e) {
            logger.error("获取Excel模板异常,原因:{}", e.getMessage());
            e.printStackTrace();
        }

        if (multiParamTemplate) {
            for (int i = 0; i < paramSheetNum; i++) {
                List sObject = new ArrayList();
                List pObject = new ArrayList();
                List fObject = new ArrayList();
                List vObject = new ArrayList();
                Sheet sheet = workbook.getSheet(i + this.skipSheets);
                if (ExcelSupport.isNotEmpty(sheet)) {
                    setParams(sheet, sObject, pObject, fObject, vObject);
                } else {
                    logger.debug("模板工作表对象不能为空!");
                }
                HashMap<String, List> map = new HashMap<String, List>();
                map.put("staticObject", sObject);
                map.put("parameterObject", pObject);
                map.put("fieldObject", fObject);
                map.put("variableObject", vObject);
                mObjectList.add(map);
            }
        } else {
            staticObject = new ArrayList();
            parameterObject = new ArrayList();
            fieldObject = new ArrayList();
            variableObject = new ArrayList();
            Sheet sheet = workbook.getSheet(this.skipSheets);
            if (ExcelSupport.isNotEmpty(sheet)) {
                setParams(sheet, staticObject, parameterObject, fieldObject, variableObject);
            } else {
                logger.debug("模板工作表对象不能为空!");
            }
        }
        workbook.close();
    }

    private void setParams(final Sheet sheet, final List staticObject, final List parameterObject, final List fieldObject, final List variableObject) {
        int rows = sheet.getRows();
        for (int k = 0; k < rows; k++) {
            Cell[] cells = sheet.getRow(k);
            for (int j = 0; j < cells.length; j++) {
                String cellContent = cells[j].getContents().trim();
                if (!ExcelSupport.isEmpty(cellContent)) {
                    if (cellContent.indexOf("$P") != -1 || cellContent.indexOf("$p") != -1) {
                        parameterObject.add(cells[j]);
                    } else if (cellContent.indexOf("$F") != -1 || cellContent.indexOf("$f") != -1) {
                        fieldObject.add(cells[j]);
                    } else if (cellContent.indexOf("$V") != -1 || cellContent.indexOf("$v") != -1) {
                        variableObject.add(cells[j]);
                    } else {
                        staticObject.add(cells[j]);
                    }
                }
            }
        }
    }

    /**
     * 增加一个静态文本对象
     */
    public void addStaticObject(Cell cell) {
        staticObject.add(cell);
    }

    /**
     * 增加一个参数对象
     */
    public void addParameterObject(Cell cell) {
        parameterObject.add(cell);
    }

    /**
     * 增加一个字段对象
     */
    public void addFieldObject(Cell cell) {
        fieldObject.add(cell);
    }

    public List getStaticObject(int i) {
        if (this.mObjectList.size() > i) {
            staticObject = this.mObjectList.get(i).get("staticObject");
        }
        return staticObject;
    }

    public List getParameterObject(int i) {
        if (this.mObjectList.size() > i) {
            parameterObject = this.mObjectList.get(i).get("parameterObject");
        }
        return parameterObject;
    }

    public List getFieldObject(int i) {
        if (this.mObjectList.size() > i) {
            fieldObject = this.mObjectList.get(i).get("fieldObject");
        }
        return fieldObject;
    }

    public List getVariableObject(int i) {
        if (this.mObjectList.size() > i) {
            variableObject = this.mObjectList.get(i).get("variableObject");
        }
        return variableObject;
    }
}
