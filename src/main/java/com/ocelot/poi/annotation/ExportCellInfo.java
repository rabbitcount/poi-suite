package com.ocelot.poi.annotation;

import java.lang.annotation.*;

/**
 * @description:
 * @author: 盗梦的虎猫
 * @date: 2018/6/28
 * @time: 6:26 PM
 * Copyright (C) 2018 AnOcelot
 * All rights reserved
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD})
public @interface ExportCellInfo {

    /**
     * 索引顺序
     *
     * @return
     */
    int index();

    /**
     * excel 列名
     *
     * @return
     */
    String columnName();

    /**
     * 当前针对日期格式
     *
     * @return
     */
    String format() default "";

}

