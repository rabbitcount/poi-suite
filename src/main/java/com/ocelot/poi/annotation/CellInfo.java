package com.ocelot.poi.annotation;

import java.lang.annotation.*;

/**
 * @description:
 * @author: 盗梦的虎猫
 * @date: 2018/6/28
 * @time: 6:27 PM
 * Copyright (C) 2018 AnOcelot
 * All rights reserved
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface CellInfo {

    String value();

    String format() default "";

    boolean isPrimary() default false;

    boolean notNull() default true;

    int length() default Integer.MAX_VALUE;
}
