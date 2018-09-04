package com.yeahpi.util;


import org.apache.commons.lang3.exception.ExceptionUtils;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ReflectUtil {

    private ReflectUtil(){}
    /**
     * 获取某个对象的某一个属性
     * @param obj
     * @param fieldName
     * @return
     */
    public static Field getFieldName(Object obj, String fieldName) {
        for (Class<?> superClass = obj.getClass(); superClass != Object.class; superClass = superClass.getSuperclass()) {
            try {
                return superClass.getDeclaredField(fieldName);
            } catch (NoSuchFieldException e) {
                System.out.println("ReflectUtil getFieldName error:"+ ExceptionUtils.getStackTrace(e));
            }
        }
        return null;
    }

    /**
     * 设置对象的属性值
     * @param obj
     * @param fieldName
     * @param value
     */
    public static void setProperty(Object obj, String fieldName, Object value) {
        try {
            Field field = obj.getClass().getDeclaredField(fieldName);
            if (field != null) {
                Class<?> fieldType = field.getType();
                field.setAccessible(true);
                //根据字段类型给字段赋值
                if (String.class == fieldType) {
                    field.set(obj, String.valueOf(value));
                } else if (Integer.class == fieldType) {
                    field.set(obj, Integer.valueOf(value.toString()));
                } else if (Long.class == fieldType) {
                    field.set(obj, Long.valueOf(value.toString()));
                } else if (Float.class == fieldType) {
                    field.set(obj, Float.valueOf(value.toString()));
                } else if (Double.class == fieldType) {
                    field.set(obj, Double.valueOf(value.toString()));
                } else if (Short.class == fieldType) {
                    field.set(obj, Short.valueOf(value.toString()));
                } else if (Character.TYPE == fieldType) {
                    if ((value != null) && (value.toString().length() > 0)) {
                        field.set(obj, Character.valueOf(value.toString().charAt(0)));
                    }
                } else if (Date.class == fieldType) {
                    if (value instanceof Date) {
                        field.set(obj, value);
                    } else if (value instanceof String) {
                        field.set(obj, new SimpleDateFormat(
                                "yyyy-MM-dd HH:mm:ss").parse(value.toString()));
                    }
                } else {
                    field.set(obj, value);
                }
                field.setAccessible(false);
            }
        }catch (Exception e) {
            e.getStackTrace();
        }
    }

    /**
     * 获取对象属性值
     * @param obj
     * @param fieldName
     * @return
     */
    public static Object getProperty(Object obj, String fieldName) {
        Field field = getFieldName(obj, fieldName);
        Object value = null;
        try {
            if (field != null) {
                field.setAccessible(true);
                value = field.get(obj);
                field.setAccessible(false);
            }
        } catch (Exception e) {
            System.out.println("ReflectUtil getProperty error:"+ ExceptionUtils.getStackTrace(e));
        }
        return value;
    }

}
