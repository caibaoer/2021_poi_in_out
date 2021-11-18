package com.hp.anotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Classname MineExcelAnnotation
 * @Description 注解
 * @Date 2021/11/17 13:51
 * @Created by huangwencai
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface MineExcelHeaderAnnotation {
    //列名
    String cloHeader();
    //列顺序
    int index() default 0;
}
