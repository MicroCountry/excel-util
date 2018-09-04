package com.yeahpi.handler;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Class
 *
 * @author wgm
 * @date 2018/09/04
 */
public class TimeDateColumnHandler implements ColumnHandler{
    /**
     * 传入时间戳
     * @param object
     * @return
     */
    @Override
    public Cell cell(Cell cell, Object object) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String value = simpleDateFormat.format(new Date(Long.valueOf(String.valueOf(object))));
        cell.setCellValue(value);
        return cell;
    }
}
