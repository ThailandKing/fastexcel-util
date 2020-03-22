# 基于FastExcel工具类的封装

## 零、实现的功能

- 基于注解化的导出配置，实现单Sheet页、多Sheet页导出excel文件
- 基于注解化的导入配置，实现数据解析，封装到实体类中

## 一、POI

-  最流行的一个(Apache POI)包含许多特性，但当涉及到大型工作表时，它很快就会占用大量内存，爆发OOM问题
-  滑动窗口机制阻止访问当前写入位置之上的单元格
-  会将内容写入临时文件
-  默认情况下禁用了共享字符串，在文件大小上带来了开销；如果启用共享字符串可能会消耗更多堆内存 

## 二、FastExcel

- 查询大量资料，发现**FastExcel**对上述问题有着很好的解决
- 基于POI封装，对外提供了简单的API
-  通过只积累必要的元素来减少内存占用和提高性能 
-  XML内容在最后通过管道传输到输出流 
- 多线程支持，多sheet页表格快速生成
- [FastExcel详解](https://github.com/dhatim/fastexcel)

## 三、封装操作过程

### 1、声明注解

- 用于实体类字段上，导出配置

```java
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ShowStyle {
    // 列名
    String displayName() default "";

    // 列宽
    double width() default 10;

    // 导出日期格式
    String format() default "";

    // 是否为导入字段标识
    boolean isImport() default false;
}
```

### 2、单Sheet页导出API

- data：待导出的数据
- clazz：导出数据的实体类字节码
- isExportTitle：是否导出表头
- response：HTTP响应
- fileName：导出文件名称

```java
/**
         * @Author thailandking
         * @Date 2020/1/8 10:00
         * @LastEditors thailandking
         * @LastEditTime 2020/1/8 10:00
         * @Description 1、List数据导出Excel（浏览器）（单sheet页）
         */
public static void exportDataBrowserSingle(List<Object> data,
                                           Class clazz, 
                                           boolean isExportTitle, 
                                           HttpServletResponse response, 
                                           String fileName) throws Exception {

}
```

### 3、多Sheet页导出API

- dataMap：每个Sheet页对应待导出的数据
- clazz：每个Sheet页导出数据对应的实体类字节码
- isExportTitle：是否导出表头
- response：HTTP响应
- fileName：导出文件名称

```java
/**
     * @Author thailandking
     * @Date 2020/1/8 10:41
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 10:41
     * @Description 2、List数据导出Excel（浏览器）（多sheet页）
     */
public static void exportDataBrowserMultiple(Map<String, List<Object>> dataMap, 
                                             Class[] classes, 
                                             boolean isExportTitle, 
                                             HttpServletResponse response, 
                                             String fileName) throws Exception {
    
}
```

### 4、导入解析

- is：输入流
- clazz：要解析成实体的字节码文件
- isIncludeTitle：是否解析表头

```java
/**
     * @Author thailandking
     * @Date 2020/1/8 15:11
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 15:11
     * @Description 3、Excel导入解析
     */
public static <T> List<T> importDataBrowser(InputStream is, 
                                            Class<T> clazz,
                                            boolean isIncludeTitle) throws Exception {
    
}
```

## 四、导出的简单使用

### 1、添加依赖

```xml
<!--fastexcel-writer-->
<dependency>
    <groupId>org.dhatim</groupId>
    <artifactId>fastexcel</artifactId>
    <version>0.10.10</version>
</dependency>
```

### 2、导入代码

- ShowStyle.java
- ExcelUtil.java

### 3、为实体类添加注解

- 特别提醒：需要添加Lombok依赖包

```java
@Data
@AllArgsConstructor //非必需，做测试填充数据用
@NoArgsConstructor
public class User {
    @ShowStyle(displayName = "编号", isImport = true)
    private Long id;
    @ShowStyle(displayName = "姓名", width = 15, isImport = true)
    private String name;
    private String pass;
    @ShowStyle(displayName = "备注说明", width = 20, isImport = true)
    private String notes;
    @ShowStyle(displayName = "状态")
    private String status;
    @ShowStyle(displayName = "注册日期", width = 20, format = "yyyy年MM月dd日", isImport = true)
    private Date joinDate;
}
```

### 4、调用

- 单Sheet页

```java
// http://localhost:9022/test/export/single
@GetMapping(value = "/export/single")
public void exportSingle(HttpServletResponse response) {
    List<Object> data = new LinkedList<>();
    User user1 = new User(1L, "wgt", "wgt", "普通用户", "激活", new Date());
    User user2 = new User(2L, "shw", "shw", "未激活用户", "未激活", new Date());
    data.add(user1);
    data.add(user2);
    try {
        ExcelUtil.exportDataBrowserSingle(data, User.class, true, response, "人员.xlsx");
    } catch (Exception e) {
        e.printStackTrace();
    }
}
```

- 多Sheet页

```java
// http://localhost:9022/test/export/multiply
@GetMapping(value = "/export/multiply")
public void exportMultiply(HttpServletResponse response) {
    Map<String, List<Object>> dataMap = new HashMap<>();
    Class[] classes = new Class[3];
    List<Object> data1 = new LinkedList<>();
    List<Object> data2 = new LinkedList<>();
    List<Object> data3 = new LinkedList<>();
    User user1 = new User(1L, "wgt", "wgt", "普通用户", "激活", new Date());
    User user2 = new User(2L, "shw", "shw", "未激活用户", "未激活", new Date());
    data1.add(user1);
    data2.add(user2);
    data3.add(user1);
    data3.add(user2);
    dataMap.put("Sheet 1", data1);
    dataMap.put("Sheet 2", data2);
    dataMap.put("Sheet 3", data3);
    classes[0] = User.class;
    classes[1] = User.class;
    classes[2] = User.class;
    try {
        ExcelUtil.exportDataBrowserMultiple(dataMap, classes, true, response, "多人员.xlsx");
    } catch (Exception e) {
        e.printStackTrace();
    }
}
```

## 五、导入的简单使用

### 1、添加依赖

```xml
<!--fastexcel-writer-->
<dependency>
    <groupId>org.dhatim</groupId>
    <artifactId>fastexcel</artifactId>
    <version>0.10.10</version>
</dependency>
<!--fastexcel-reader-->
<dependency>
    <groupId>org.dhatim</groupId>
    <artifactId>fastexcel-reader</artifactId>
    <version>0.10.2</version>
</dependency>
```

### 2、导入代码

- ShowStyle.java
- ExcelUtil.java

### 3、为实体类添加注解

- 特别提醒：需要添加Lombok依赖包

```java
@Data
@NoArgsConstructor
public class People {
    @ShowStyle(isImport = true)
    private Integer id;
    @ShowStyle(isImport = true)
    private Short num;
    @ShowStyle(isImport = true)
    private Long count;
    @ShowStyle(isImport = true)
    private Float money;
    @ShowStyle(isImport = true)
    private Double balance;
    @ShowStyle(isImport = true)
    private Boolean flag;
    @ShowStyle(isImport = true)
    private Byte b;
    @ShowStyle(isImport = true)
    private String name;
    @ShowStyle(isImport = true)
    private Date joinDate;

    private String pass;
}
```

### 4、调用

```java
// http://localhost:9022/test/import/complex
@PostMapping("/import/complex")
public List<People> importExcelComplex(@RequestParam("file") MultipartFile file) {
    try {
        InputStream is = file.getInputStream();
        List<People> peoples = ExcelUtil.importDataBrowser(is, People.class, true);
        peoples.forEach(item -> System.out.println(item));
        return peoples;
    } catch (Exception e) {
        e.printStackTrace();
    }
    return null;
}
```

## 特别提醒

- 导出：
- 避免大数据量（业务尽量避免一次查询过多数据）
- 导入：
- 单元格样式设置为文本、日期格式默认 yyyy/MM/dd、实体类必须提供空参构造函数

