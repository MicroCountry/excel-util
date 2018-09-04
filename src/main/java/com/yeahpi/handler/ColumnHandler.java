package com.yeahpi.handler;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Class
 *
 * @author wgm
 * @date 2018/09/04
 */
public interface ColumnHandler {
    Cell cell(Cell cell, Object object);
}
