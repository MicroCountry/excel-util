## excel-util
### 将java类与excel互转

 1.将list对象导出到excel中,可以使用自定义Handler定义输出数据格式
```
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
 2.枚举类优化,这里枚举类有一定要求，属性必须为code,desc，用户可以根据自己需求修改，也可以使用例1中的资历
```$xslt
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
```