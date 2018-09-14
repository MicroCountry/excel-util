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
 3. controller导出接口
 ```$xslt
   @RequestMapping(value = "/exportShopList", method = RequestMethod.GET)
   @ResponseBody
   public void exportShopList( HttpServletResponse response) throws Exception{
       List<JHShopInfo> shopList = new ArrayList<>();
       if(shopList == null || shopList.length == 0){
           throw new AppException(ApiError.NO_DATA.getCode(), ApiError.NO_DATA.getMsg());
       }
       SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
       Map<String,String> map = new LinkedHashMap<>();
       map.put("shopName", "门店名称");
       map.put("shopType", "门店类型");
       map.put("shopStatus", "门店状态");
       map.put("shopAddress", "门店地址");
       map.put("shopContactName", "联系人");
       map.put("shopTel", "电话");
       map.put("shopMobile", "手机号码");
       map.put("shopCreateTime", "创建时间");

       Map<String , ColumnHandler> handlerMap = new HashMap<>();
       TimeDateColumnHandler dateColumnHandler = new TimeDateColumnHandler();
       handlerMap.put("shopType", new EnumColumnHandler<ShopType>(ShopType.class));
       handlerMap.put("shopStatus", new EnumColumnHandler<ShopStatus>(ShopStatus.class));
       handlerMap.put("shopCreateTime", dateColumnHandler);

       String fileName = "门店列表-"+dateFormat.format(new Date(System.currentTimeMillis())) +".xls";
       String str = path + fileName;
       File file = new File(str);
       ExcelHandleUtils.exportToExcel(file, sheetName, map, handlerMap, shopList);
       //清除buffer缓存
       response.reset();
       response.setHeader("Content-Disposition", "attachment;filename="+ URLEncoder.encode(fileName,"UTF-8"));
       response.setContentType("application/octet-stream;charset=UTF-8");
       response.setHeader("Pragma", "no-cache");
       response.setHeader("Cache-Control", "no-cache");
       response.setDateHeader("Expires", 0);

       try {
           FileInputStream fis = new FileInputStream(file);
           byte[] b = new byte[fis.available()];
           fis.read(b);
           //获取响应报文输出流对象
           ServletOutputStream out =response.getOutputStream();
           //输出
           out.write(b);
           out.flush();
           out.close();
       } catch (IOException e) {
           e.printStackTrace();
       }
   }
```