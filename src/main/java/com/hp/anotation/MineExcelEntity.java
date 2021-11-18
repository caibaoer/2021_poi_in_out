package com.hp.anotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Classname ExcelEntity
 * @Description ExcelEntity
 * @Date 2021/11/17 16:36
 * @Created by huangwencai
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface MineExcelEntity {
}
