package com.hp.entity;


import com.hp.anotation.MineExcelEntity;
import com.hp.anotation.MineExcelProperty;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @Classname ImportCaseInfo
 * @Description ImportCaseInfo 导入数据  首先要知道列顺序  列名字
 * @Date 2021/11/17 16:05
 * @Created by huangwencai
 */
@Data
@MineExcelEntity
public class ImportDemoCaseInfo {
    @MineExcelProperty(fieldName = "催收员",fieldIndex = 0,fieldType = Cell.CELL_TYPE_STRING,nullable = false)
    private String collector;
    @MineExcelProperty(fieldName = "案件号",fieldIndex = 1,fieldType = Cell.CELL_TYPE_STRING,nullable = false)
    private String projectCode;

}
