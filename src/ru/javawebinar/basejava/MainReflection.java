package ru.javawebinar.basejava;

import ru.javawebinar.basejava.model.Resume;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class MainReflection {
    public static void main(String[] args) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        Resume r = new Resume();
        Field field = r.getClass().getDeclaredFields()[0];
        field.setAccessible(true);
        System.out.println(field.getName());
        System.out.println(field.get(r));
        field.set(r, "new_uuid");
        System.out.println(r);


        Method method = r.getClass().getDeclaredMethod("toString");
        Object result = method.invoke(r);
        System.out.println(result);
        Method[] methods = r.getClass().getDeclaredMethods();
        for (Method method1 : methods) {
            System.out.println(method1.getName());
        }
    }
}
