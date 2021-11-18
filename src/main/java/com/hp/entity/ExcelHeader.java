package com.hp.entity;

import lombok.Data;

/**
 * @Classname ExcelHeader
 * @Description ExcelHeader
 * @Date 2021/11/17 14:31
 * @Created by huangwencai
 */
@Data
public class ExcelHeader implements Comparable<ExcelHeader>{
    private String colName;
    private String colProperty;
    private int index;


    @Override
    public int compareTo(ExcelHeader o) {
        return this.index-o.index;
    }
}
