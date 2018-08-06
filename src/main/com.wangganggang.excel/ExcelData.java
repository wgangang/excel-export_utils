package com.wangganggang.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.List;
import java.util.Map;

/**
 * @author wangganggang
 * @date 2018年08月05日 下午10:22
 */
@Getter
@Setter
@AllArgsConstructor
public class ExcelData {

    /**
     * Excel参数元数据对象
     */
    private Map parametersMap;

    /**
     * Excel集合元对象
     */
    private List fieldsList;
}
