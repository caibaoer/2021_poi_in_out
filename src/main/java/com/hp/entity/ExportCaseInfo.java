package com.hp.entity;

import com.hp.anotation.MineExcelHeaderAnnotation;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @Classname CaseInfo
 * @Description 导出的数据实体类  我还想要指定导出数据在哪一列的话，就需要使用到注解了。
 * @Date 2021/11/17 13:46
 * @Created by huangwencai
 */
@Data
@AllArgsConstructor
public class ExportCaseInfo {
    @MineExcelHeaderAnnotation(cloHeader="项目编号",index=0)
    private String projectCode;
    @MineExcelHeaderAnnotation(cloHeader="客户名字",index=1)
    private String customerName;
    @MineExcelHeaderAnnotation(cloHeader="催收专员",index=3)
    private String collector;
}
