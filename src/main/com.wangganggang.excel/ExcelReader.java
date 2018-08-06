package com.wangganggang.excel;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月06日 上午12:03
 */
@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class ExcelReader {

    private String metaData = null;
    private InputStream is = null;

    /**
     * 读取Excel数据
     *
     * @param pBegin 从第几行开始读数据<br>
     *               <b>注意下标索引从0开始的哦!
     * @return 以List<BaseDTO>形式返回数据
     * @throws BiffException
     * @throws IOException
     */
    public List read(int pBegin) throws BiffException, IOException {
        List list = new ArrayList();
        Workbook workbook = Workbook.getWorkbook(getIs());
        Sheet sheet = workbook.getSheet(0);
        int rows = sheet.getRows();
        for (int i = pBegin; i < rows; i++) {
            Map rowDto = new HashMap<>();
            Cell[] cells = sheet.getRow(i);
            for (int j = 0; j < cells.length; j++) {
                String key = getMetaData().trim().split(",")[j];
                if (ExcelSupport.isNotEmpty(key)) {
                    rowDto.put(key, cells[j].getContents());
                }
            }
            list.add(rowDto);
        }
        return list;
    }

    /**
     * 读取Excel数据
     *
     * @param pBegin 从第几行开始读数据<br>
     *               <b>注意下标索引从0开始的哦!</b>
     * @param pBack  工作表末尾减去的行数
     * @return 以List<BaseDTO>形式返回数据
     * @throws BiffException
     * @throws IOException
     */
    public List read(int pBegin, int pBack) throws BiffException, IOException {
        List list = new ArrayList();
        Workbook workbook = Workbook.getWorkbook(getIs());
        Sheet sheet = workbook.getSheet(0);
        int rows = sheet.getRows();
        for (int i = pBegin; i < rows - pBack; i++) {
            Map rowDto = new HashMap<>();
            Cell[] cells = sheet.getRow(i);
            String[] arrMeta = getMetaData().trim().split(",");
            for (int j = 0; j < arrMeta.length; j++) {
                String key = arrMeta[j];
                if (ExcelSupport.isNotEmpty(key)) {
                    rowDto.put(key, cells[j].getContents());
                }
            }
            list.add(rowDto);
        }
        return list;
    }
}
