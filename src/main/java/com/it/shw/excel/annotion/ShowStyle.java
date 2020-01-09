package com.it.shw.excel.annotion;

import java.lang.annotation.*;

/**
 * @Copyright: Harbin Institute of Technology.All rights reserved.
 * @Description:
 * @author: thailandking
 * @since: 2020/1/7 14:52
 * @history: 1.2020/1/7 created by thailandking
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ShowStyle {
    // 列名
    String displayName() default "";

    // 列宽
    double width() default 10;

    // 导出日期格式
    String format() default "";

    // 是否为导入字段标识
    boolean isImport() default false;
}