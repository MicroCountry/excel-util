package com.yeahpi.test;

import com.yeahpi.handler.ColumnHandler;
import com.yeahpi.handler.EnumColumnHandler;
import com.yeahpi.handler.TestEnumColumnHandler;
import com.yeahpi.handler.TimeDateColumnHandler;
import com.yeahpi.model.TestEnum;
import com.yeahpi.util.ExcelHandleUtils;

import java.io.File;
import java.util.*;

/**
 * Class
 *
 * @author wgm
 * @date 2018/09/04
 */
public class TestExport {
    public static void main(String[] args) {
        List<TestBean> list = new ArrayList<>();
        TestBean b1 = new TestBean();
        b1.setCode(1);
        b1.setId(1);
        b1.setName("测试");
        b1.setTime(System.currentTimeMillis());
        list.add(b1);

        Map<String, String> colMap = new LinkedHashMap<>();
        colMap.put("id", "id值");
        colMap.put("code", "描述");
        colMap.put("name", "姓名");
        colMap.put("time", "时间");

        Map<String , ColumnHandler> handlerMap = new HashMap<>();
        TimeDateColumnHandler dateColumnHandler = new TimeDateColumnHandler();
        handlerMap.put("code", new EnumColumnHandler<TestEnum>(TestEnum.class));
        handlerMap.put("time", dateColumnHandler);

        File file = new File("/Users/apple/workspace/excel-util/test2.xls");
        ExcelHandleUtils.exportToExcel(file, "测试", colMap, handlerMap, list);
    }
}
