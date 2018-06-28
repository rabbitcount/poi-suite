package com.ocelot.poi.annotation;

import java.lang.annotation.*;

/**
 * @description:
 * @author: 盗梦的虎猫
 * @date: 2018/6/28
 * @time: 6:28 PM
 * Copyright (C) 2018 AnOcelot
 * All rights reserved
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface CellListInfo {

    String regex();

    int groupName();

    int group();

    int columnNameGroup();
}
