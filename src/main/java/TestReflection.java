import lombok.*;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * @description:
 * @author: gaoxiang09
 * @date: 2018/6/28
 * @time: 7:37 PM
 * Copyright (C) 2018 Meituan
 * All rights reserved
 */
public class TestReflection {

    public static void main(String[] args) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        List<Entity> source = new ArrayList<>();
        Method method = Collection.class.getDeclaredMethod("add", Object.class);
        for (int index = 0; index < 10; ++index) {
            Entity e = new Entity(index, "test" + index);
            method.invoke(source, e);
        }
        for (Entity e : source) {
            System.out.println(e);
        }
    }

    @Setter
    @Getter
    @NoArgsConstructor
    @AllArgsConstructor
    @ToString
    public static class Entity {
        int index;
        String message;
    }
}
