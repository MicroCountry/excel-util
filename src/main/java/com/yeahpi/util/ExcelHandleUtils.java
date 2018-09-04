package com.yeahpi.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.yeahpi.handler.ColumnHandler;
import jxl.CellType;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.WritableCellFormat;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelHandleUtils { 

    private static SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    /**
     * 把 excel中的每一个sheetName页的内容都到JSONArray中，
     * excel 表格的第一行内容为key值，第二行开始为value
     * @param in excel 文件的输入流
     * @return jsonObject 中的每一个key为sheetName
     */
    public static JSONObject readExcelToJsonObject(InputStream in, String fileName){
        if (in == null){
            return null;
        }
        Workbook workbook = getWorkboot(in, fileName);
        if (workbook != null) {
            JSONObject response = new JSONObject();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet != null) {
                    String name = sheet.getSheetName();
                    response.put(name, getSheetContents(sheet));
                }
            }
            return response;
        } else {
            System.out.println("Workboot is null");
        }
        return null;
    }

    /**
     * 把 excel 内容读成对应类型数组，请求参数不能为null，否则返回null
     * excel 表格的第一行内容为对应的class的属性名要一致
     * @param in excel 文件的输入流
     * @param fileName 文件名
     * @param sheetName 需要读取的sheet页的名称
     * @param className 转换的对象类型
     * @return
     */
    public static List readExcelToArray(InputStream in,String fileName, String sheetName, Class className) {
        if (StringUtils.isBlank(sheetName) || className == null || in == null) {
            System.out.println("read excel to array request param can not null.");
            return null;
        }
        try{
            JSONArray jsonArray = readExcelToJSONArray(in, sheetName,fileName);
            if (jsonArray != null){
                return JSON.parseArray(jsonArray.toJSONString(), className);
            }
        }catch (Exception e){
            System.out.println("Sheet contents can not parse to " + className + " type.");
            return null;
        }
        return null;
    }

    /**
     * 把 excel中 sheetName 页的内容都到JSONArray中，
     * excel 表格的第一行内容为key值，第二行开始为value
     * @param in excel 文件的输入流
     * @param sheetName 需要读取的sheet页的名称
     * @return
     */
    public static JSONArray readExcelToJSONArray(InputStream in, String sheetName, String fileName) {
        if (StringUtils.isBlank(sheetName) || in == null) {
            System.out.println("read excel to array request param can not null.");
            return null;
        }
        Workbook workbook = getWorkboot(in, fileName);
        if (workbook != null) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet != null) {
                    String name = sheet.getSheetName();
                    if (StringUtils.isNotBlank(sheetName) && sheetName.equalsIgnoreCase(name)) {
                        return getSheetContents(sheet);
                    }
                }
            }
            System.out.println("Can not find the " + sheetName + " sheet.");
        } else {
            System.out.println("Workboot is null");
        }
        return null;
    }

    private static JSONArray getSheetContents(Sheet sheet) {
        JSONArray jsonArray = new JSONArray();
        if (sheet != null) {
            List<String> headList = new ArrayList<>();
            int phyRow = sheet.getPhysicalNumberOfRows();
            for (int i = 0; i < phyRow; i++) {
                if (i == 0) {
                    //第一行作为key值
                    Row row = sheet.getRow(i);
                    Iterator<Cell> iterator = row.cellIterator();
                    while (iterator.hasNext()){
                        Cell cell = iterator.next();
                        headList.add(cell.getStringCellValue());
                    }
                    continue;
                }
                //从第二行开始作为value的值
                JSONObject cellJson = new JSONObject();
                Row row = sheet.getRow(i);
                Iterator<Cell> iterator = row.cellIterator();
                int j =0;
                while (iterator.hasNext()){
                    Cell cell = iterator.next();
                    cellJson.put(headList.get(j), getCellValue(cell));
                    j++;
                }
                jsonArray.add(cellJson);
            }
        }
        return jsonArray;
    }


    private static Object getCellValue(Cell cell) {
        if (cell != null) {
            if (CellType.BOOLEAN.equals(cell.getCellType())) {
                return cell.getBooleanCellValue();
            } else if (CellType.DATE.equals(cell.getCellType())) {
                return cell.getDateCellValue();
            } else if (CellType.NUMBER.equals(cell.getCellType())) {
                return cell.getNumericCellValue();
            } else {
                return cell.getStringCellValue();
            }

        }
        return null;
    }

    private static Workbook getWorkboot(InputStream in, String fileName){
        Workbook workbook ;
        String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
                .substring(fileName.lastIndexOf(".") + 1);
        if ("xls".equals(extension)) {
            try {
                workbook = new HSSFWorkbook(in);
            } catch (IOException e) {
                return null;
            }
        } else if ("xlsx".equals(extension)) {
            try {
                workbook = new XSSFWorkbook(in);
            } catch (IOException e) {
                return null;
            }
        } else {
            return null;
        }
        return workbook;
    }

    /**
     * 把List的内容放入到excel文件中
     * @param file 生产excel的文件
     * @param sheetName 页签名称
     * @param headMap 第一行的头名称对应字段,
     *                key是List对象中的属性名，value是该属性显示的标题头名称，
     *                建议用LinkHashMap类型的map, 这样可以控制map的顺序
     * @param list 存放内容的List
     * @param <T> list 对应的类型
     */
    public static <T> void exportToExcelIos( File file, String sheetName, Map<String, String> headMap, Map<String, ColumnHandler> handlerMap,  List<T> list) {
        if (list == null || list.size() <= 0 || file == null) {
            System.out.println("存入excel的List内容不存在或文件不存在！");
            return;
        }
        if (headMap == null) {
            System.out.println("需要存入excel的内容没有对应的标识头,必输输入标识头");
            return;
        }

        if (StringUtils.isBlank(sheetName)) {
            sheetName = "sheet";
        }
        try {
            String fileName = file.getName();
            String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
                    .substring(fileName.lastIndexOf(".") + 1);

            Workbook wb = null;
            if ("xls".equals(extension)) {
                try {
                    wb = new HSSFWorkbook();
                } catch (Exception e) {
                    System.out.println("error");
                }
            } else if ("xlsx".equals(extension)) {
                try {
                    wb = new SXSSFWorkbook(100);
                } catch (Exception e) {
                    System.out.println("error");
                }
            }
            Sheet sheet = wb.createSheet(sheetName);
            fillSheet(sheet, headMap, handlerMap, list, 0, list.size()); //填充内容
            FileOutputStream fOut = new FileOutputStream(file);
            wb.write(fOut);
            //刷新缓冲区
            fOut.flush();
            fOut.close();
        }catch (Exception e){
            System.out.println("[exportToExcel] error=");
        }

    }

    /**
     * 把List的内容放入到excel文件中
     * @param file 生产excel的文件
     * @param sheetName 页签名称
     * @param headMap 第一行的头名称对应字段,
     *                key是List对象中的属性名，value是该属性显示的标题头名称，
     *                建议用LinkHashMap类型的map, 这样可以控制map的顺序
     * @param list 存放内容的List
     * @param <T> list 对应的类型
     */
    public static <T> void exportToExcel(File file, String sheetName, Map<String, String> headMap, Map<String, ColumnHandler> handlerMap, List<T> list) {
        if (list == null || list.size() <= 0 || file == null) {
            System.out.println("存入excel的List内容不存在或文件不存在！");
            return;
        }
        if (headMap == null) {
            System.out.println("需要存入excel的内容没有对应的标识头,必输输入标识头");
            return;
        }

        if (StringUtils.isBlank(sheetName)) {
            sheetName = "sheet";
        }
        try {
            String fileName = file.getName();
            String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
                    .substring(fileName.lastIndexOf(".") + 1);

            Workbook wb = null;
            if ("xls".equals(extension)) {
                try {
                    wb = new HSSFWorkbook();
                } catch (Exception e) {
                    System.out.println("error");
                }
            } else if ("xlsx".equals(extension)) {
                try {
                    wb = new SXSSFWorkbook(100);
                } catch (Exception e) {
                    System.out.println("error");
                }
            }
            Sheet sheet = wb.createSheet(sheetName);
            //填充内容
            fillSheet(sheet, headMap, handlerMap, list,0, list.size());
            FileOutputStream fOut = new FileOutputStream(file);
            wb.write(fOut);
            //刷新缓冲区
            fOut.flush();
            fOut.close();
        }catch (Exception e){
            System.out.println("[exportToExcel] error=");
        }

    }

    private static <T> void fillSheet(Sheet sheet, Map<String, String> headMap, Map<String, ColumnHandler> handlerMap, List<T> list, int startIndex, int size) throws Exception {
        //定义存放T的字段名和描述名的数组
        String[] fieldNameArr = new String[headMap.size()];
        String[] fieldDescArr = new String[headMap.size()];
        int fieldCount = 0;
        for (Map.Entry<String, String> entry : headMap.entrySet()) {
            fieldNameArr[fieldCount] = entry.getKey();
            fieldDescArr[fieldCount] = entry.getValue();
            fieldCount ++;
        }

        //设置表头样式
        WritableCellFormat headCF = new WritableCellFormat();
        //单元格颜色
        headCF.setBackground(Colour.GRAY_25);
        //单元格居中
        headCF.setAlignment(Alignment.CENTRE);

        //填充表头
        Row row1 = sheet.createRow(0);
        for (int i = 0; i < fieldDescArr.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(fieldDescArr[i]);
        }

        //填充内容
        //从第二行开始填list的内容
        int rowNo = 1;
        for (int index = startIndex; index < size; index ++) {
            Row row = sheet.createRow(rowNo);
            //获取单个对象
            T item = list.get(index);
            for (int colNo = 0; colNo < fieldNameArr.length; colNo ++) {
                //通过反射获取到对应的属性值
                Object value = ReflectUtil.getProperty(item, fieldNameArr[colNo]);
                Cell cell = row.createCell(colNo);
                if(handlerMap.get(fieldNameArr[colNo]) != null){
                    ColumnHandler columnHandler = handlerMap.get(fieldNameArr[colNo]);
                    columnHandler.cell(cell, value);
                }else {
                    String fieldValue = "";
                    if (value != null && value instanceof Date) {
                        //如果是时间类型的数据，直接格式换为字符串
                        fieldValue = df1.format(value);
                    } else if (value != null) {
                        fieldValue = value.toString();
                    }
                    cell.setCellValue(fieldValue);
                }
            }
            rowNo ++;
        }
    }

}
