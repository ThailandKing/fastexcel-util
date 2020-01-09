package com.it.shw.excel.utils;

import com.it.shw.excel.annotion.ShowStyle;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;

/**
 * @Copyright: Harbin Institute of Technology.All rights reserved.
 * @Description:
 * @author: thailandking
 * @since: 2020/1/7 14:35
 * @history: 1.2020/1/7 created by thailandking
 */
public class ExcelUtil {

    // TODO
    // 导出
    // 1、时间格式化 over
    // 2、filename over
    // 3、多个sheet页数据 over
    // 4、表头选择   over
    // 5、大数据量（业务问题，尽量避免一次查询过多数据）

    // 导入
    // 1、单元格样式设置为文本
    // 2、表头选择
    // 3、日期格式默认 yyyy/MM/dd

    /**
     * @Author thailandking
     * @Date 2020/1/8 9:22
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 9:22
     * @Description 设置表头（列名）
     */
    private static void setExportTitle(Worksheet ws, Field[] fields, boolean isExportTitle) {
        int titleRow = 0;
        int titleColumn = 0;
        for (Field field : fields) {
            ShowStyle style = field.getAnnotation(ShowStyle.class);
            if (style != null) {
                String displayName = style.displayName();
                double width = style.width();
                ws.width(titleColumn, width);
                // 样式保留、内容去掉
                if (isExportTitle) {
                    ws.value(titleRow, titleColumn, displayName);
                }
                titleColumn++;
            }
        }
    }

    /**
     * @Author thailandking
     * @Date 2020/1/8 10:45
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 10:45
     * @Description 设置sheet页内容
     */
    private static void setExportSheet(Workbook wb, List<Object> data, Class clazz, boolean isExportTitle, String sheetName) throws Exception {
        Worksheet ws = wb.newWorksheet(sheetName);
        Field[] fields = clazz.getDeclaredFields();
        // 设置表头
        setExportTitle(ws, fields, isExportTitle);
        // 填充表格内容
        int bodyRow = (isExportTitle == true ? 1 : 0);
        int bodyColumn = 0;
        for (int i = 0; i < data.size(); i++) {
            Object item = data.get(i);
            for (Field field : fields) {
                ShowStyle style = field.getAnnotation(ShowStyle.class);
                if (style != null) {
                    // 设置私有为公开
                    field.setAccessible(true);
                    Object value = field.get(item);
                    // 格式化日期数据
                    String format = style.format();
                    if (format != null && !format.isEmpty()) {
                        SimpleDateFormat sdf = new SimpleDateFormat(format);
                        value = sdf.format(value);
                    }
                    ws.value(bodyRow, bodyColumn, value);
                    bodyColumn++;
                }
            }
            bodyRow++;
            bodyColumn = 0;
        }
    }

    /**
     * @Author thailandking
     * @Date 2020/1/8 11:00
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 11:00
     * @Description HttpServletResponse获取输出流
     */
    private static ServletOutputStream getOutputStream(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.addHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF8"));
        ServletOutputStream os = response.getOutputStream();
        return os;
    }

    /**
     * @Author thailandking
     * @Date 2020/1/8 9:24
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 9:24
     * @Description List数据导出Excel（单sheet页）
     */
    private static void exportListData(List<Object> data, Class clazz, boolean isExportTitle, OutputStream os) throws Exception {
        if (data == null || data.isEmpty()) {
            return;
        }
        // 创建excel对象
        Workbook wb = new Workbook(os, "MyApplication", "1.0");
        setExportSheet(wb, data, clazz, isExportTitle, "Sheet 1");
        // 输出流
        wb.finish();
    }

    /**
     * @Author thailandking
     * @Date 2020/1/8 11:04
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 11:04
     * @Description List数据导出Excel（多sheet页）
     */
    private static void exportSheetData(Map<String, List<Object>> dataMap, Class[] classes, boolean isExportTitle, OutputStream os) throws Exception {
        if (dataMap == null || dataMap.isEmpty()) {
            return;
        }
        // 创建excel对象
        Workbook wb = new Workbook(os, "MyApplication", "1.0");
        // 多线程填充sheet
        CompletableFuture<Void>[] futures = new CompletableFuture[dataMap.size()];
        int index = 0;
        for (Map.Entry<String, List<Object>> entry : dataMap.entrySet()) {
            int pos = index;
            CompletableFuture<Void> cf = CompletableFuture.runAsync(() -> {
                try {
                    setExportSheet(wb, entry.getValue(), classes[pos], isExportTitle, entry.getKey());
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            });
            futures[index] = cf;
            index++;
        }
        // 多线程执行
        CompletableFuture.allOf(futures).get();
        wb.finish();
    }

    /**
     * @Author thailandking
     * @Date 2020/1/9 9:02
     * @LastEditors thailandking
     * @LastEditTime 2020/1/9 9:02
     * @Description 将Excel单元格数据转换为Java字段值
     */
    private static void getCellValue(Row row, int index, Field field, Object o) throws Exception {
        field.setAccessible(true);
        Cell cell = row.getCell(index);
        // Integer
        if (field.getType().isAssignableFrom(Integer.class) || field.getType().getName().equals("int")) {
            field.set(o, Integer.valueOf(cell.getValue().toString()));
        }
        // Short
        else if (field.getType().isAssignableFrom(Short.class) || field.getType().getName().equals("short")) {
            field.set(o, Short.valueOf(cell.getValue().toString()));
        }
        // Float
        else if (field.getType().isAssignableFrom(Float.class) || field.getType().getName().equals("float")) {
            field.set(o, Float.valueOf(cell.getValue().toString()));
        }
        // Byte
        else if (field.getType().isAssignableFrom(Byte.class) || field.getType().getName().equals("byte")) {
            field.set(o, Byte.valueOf(cell.getValue().toString()));
        }
        // Double
        else if (field.getType().isAssignableFrom(Double.class) || field.getType().getName().equals("double")) {
            field.set(o, Double.valueOf(cell.getValue().toString()));
        }
        // Long
        else if (field.getType().isAssignableFrom(Long.class) || field.getType().getName().equals("long")) {
            field.set(o, Long.valueOf(cell.getValue().toString()));
        }
        // Boolean
        else if (field.getType().isAssignableFrom(Boolean.class) || field.getType().getName().equals("boolean")) {
            field.set(o, Boolean.valueOf(cell.getValue().toString()));
        }
        // String
        else if (field.getType().isAssignableFrom(String.class)) {
            field.set(o, cell.getValue().toString());
        }
        // Date  默认yyyy/MM/dd
        else if (field.getType().isAssignableFrom(Date.class)) {
            String value = cell.getValue().toString();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
            Date date = sdf.parse(value);
            field.set(o, date);
        }
    }

    /**
     * @Author thailandking
     * @Date 2020/1/8 10:00
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 10:00
     * @Description 1、List数据导出Excel（浏览器）（单sheet页）
     */
    public static void exportDataBrowserSingle(List<Object> data, Class clazz, boolean isExportTitle, HttpServletResponse response, String fileName) throws Exception {
        exportListData(data, clazz, isExportTitle, getOutputStream(response, fileName));
    }

    /**
     * @Author thailandking
     * @Date 2020/1/8 10:41
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 10:41
     * @Description 2、List数据导出Excel（浏览器）（多sheet页）
     */
    public static void exportDataBrowserMultiple(Map<String, List<Object>> dataMap, Class[] classes, boolean isExportTitle, HttpServletResponse response, String fileName) throws Exception {
        exportSheetData(dataMap, classes, isExportTitle, getOutputStream(response, fileName));
    }


    // 一次读入缓存，OOM问题（通过stream流封装）
    //try (InputStream is = ...; ReadableWorkbook wb = new ReadableWorkbook(is)) {
    //    Sheet sheet = wb.getFirstSheet();
    //    try (Stream<Row> rows = sheet.openStream()) {
    //        rows.forEach(r -> {
    //            BigDecimal num = r.getCellAsNumber(0).orElse(null);
    //            String str = r.getCellAsString(1).orElse(null);
    //            LocalDateTime date = r.getCellAsDate(2).orElse(null);
    //        });
    //    }
    //}

    /**
     * @Author thailandking
     * @Date 2020/1/8 15:11
     * @LastEditors thailandking
     * @LastEditTime 2020/1/8 15:11
     * @Description 3、Excel导入解析
     */
    public static <T> List<T> importDataBrowser(InputStream is, Class<T> clazz, boolean isIncludeTitle) throws Exception {
        // 创建excel对象
        ReadableWorkbook wb = new ReadableWorkbook(is);
        Sheet sheet = wb.getFirstSheet();
        // 一次读入缓存，OOM问题（通过stream流封装）
        List<Row> rows = sheet.read();
        if (rows == null || rows.isEmpty()) {
            return new ArrayList<>();
        }
        // 解析数据
        List<T> returnData = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        int start = (isIncludeTitle == true ? 1 : 0);
        // 遍历每一行
        for (int i = start; i < rows.size(); i++) {
            Row row = rows.get(i);
            T t = clazz.newInstance();
            boolean flag = false;
            // 只为支持导入的字段赋值
            int index = 0;
            for (Field field : fields) {
                ShowStyle style = field.getAnnotation(ShowStyle.class);
                if (style != null) {
                    boolean isImport = style.isImport();
                    if (isImport) {
                        // 通过反射机制赋值
                        getCellValue(row, index, field, t);
                        index++;
                        flag = true;
                    }
                }
            }
            // 如果没有支持导入字段、不添加空对象
            if (flag) {
                returnData.add(t);
            }
        }
        return returnData;
    }
}

