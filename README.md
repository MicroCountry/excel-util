## excel-util
### 将java类与excel互转

```
    1.将list对象导出到excel中,可以使用自定义Handler定义输出数据格式
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
        TestEnumColumnHandler enumColumnHandler = new TestEnumColumnHandler();
        TimeDateColumnHandler dateColumnHandler = new TimeDateColumnHandler();
        handlerMap.put("code", enumColumnHandler);
        handlerMap.put("time", dateColumnHandler);
    
        File file = new File("/Users/apple/workspace/excel-util/test1.xls");
        ExcelHandleUtils.exportToExcel(file, "测试", colMap, handlerMap, list);
```