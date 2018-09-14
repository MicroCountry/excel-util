package com.yeahpi.handler;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Method;

/**
 * Class
 *
 * @author wgm
 * @date 2018/09/14
 */
public class EnumColumnHandler<E extends Enum>  implements ColumnHandler {

    private Class<E> type;

    @Override
    public Cell cell(Cell cell, Object object) {
        int code = (int) object;
        String desc =getDescByCode(code);
        cell.setCellValue(desc);
        return cell;
    }

    public EnumColumnHandler(Class<E> type){
        this.type = type;
    }

    public String getDescByCode(int code){
        try {
            Method getCode = type.getMethod("getCode");
            Method getDesc = type.getMethod("getDesc");
            E[] enums = type.getEnumConstants();
            if (enums == null) {
                throw new IllegalArgumentException(type.getSimpleName() + " does not represent an enum type.");
            }
            for (E e : enums) {
                int eCode = (Integer) getCode.invoke(e);
                if(eCode == code){
                    return (String)getDesc.invoke(e);
                }
            }
        }catch (Exception e){
            System.out.println("getDescByCode error " + ExceptionUtils.getStackTrace(e));
        }
        return "";
    }


}

