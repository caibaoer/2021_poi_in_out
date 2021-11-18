package com.hp.anotation;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Classname MineExcelProperty
 * @Description 模板属性注解
 * @Date 2021/11/17 16:11
 * @Created by huangwencai
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
public @interface MineExcelProperty {
    //列名
    String fieldName() default "";
    //EXCEL 列号
    int fieldIndex() default -1;

    /**
     *   EXCEl CELL 类型  0数字   1字符串 2公式 3空值 4 Boolea值  5非法字符
     *
     *   0、Cell.CELL_TYPE_NUMERIC
     *   1、Cell.CELL_TYPE_STRING
     *   2、Cell.CELL_TYPE_FORMULA
     *   3、Cell.CELL_TYPE_BLANK
     *   4、Cell.CELL_TYPE_BOOLEAN
     *   5、Cell.CELL_TYPE_ERROR
     */

    int fieldType() default -1;
    //是否可以为空
    boolean nullable() default false;
    //最大长度
    int maxLength() default 0;
    //最小长度
    int minLength() default 0;
}
