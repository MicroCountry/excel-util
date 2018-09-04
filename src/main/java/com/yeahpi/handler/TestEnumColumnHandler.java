package com.yeahpi.handler;

import com.yeahpi.model.TestEnum;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Class
 *
 * @author wgm
 * @date 2018/09/04
 */
public class TestEnumColumnHandler implements ColumnHandler {

    @Override
    public Cell cell(Cell cell, Object object) {
        int code = (int) object;
        String desc = TestEnum.getDescByCode(code);
        cell.setCellValue(desc);
        return cell;
    }
}
