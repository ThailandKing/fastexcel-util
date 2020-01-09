package com.it.shw.excel.controller;

import com.it.shw.excel.dto.People;
import com.it.shw.excel.dto.User;
import com.it.shw.excel.utils.ExcelUtil;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.*;

/**
 * @Copyright: Harbin Institute of Technology.All rights reserved.
 * @Description:
 * @author: thailandking
 * @since: 2020/1/8 9:58
 * @history: 1.2020/1/8 created by thailandking
 */
@RestController
@RequestMapping(value = "/test")
public class TestController {

    // http://localhost:9022/test/export/single
    @GetMapping(value = "/export/single")
    public void exportSingle(HttpServletResponse response) {
        List<Object> data = new LinkedList<>();
        User user1 = new User(1L, "王国泰", "wgt", "普通用户", "激活", new Date());
        User user2 = new User(2L, "shw", "shw", "未激活用户", "未激活", new Date());
        data.add(user1);
        data.add(user2);
        try {
            ExcelUtil.exportDataBrowserSingle(data, User.class, true, response, "人员.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // http://localhost:9022/test/export/multiply
    @GetMapping(value = "/export/multiply")
    public void exportMultiply(HttpServletResponse response) {
        Map<String, List<Object>> dataMap = new HashMap<>();
        Class[] classes = new Class[3];
        List<Object> data1 = new LinkedList<>();
        List<Object> data2 = new LinkedList<>();
        List<Object> data3 = new LinkedList<>();
        User user1 = new User(1L, "王国泰", "wgt", "普通用户", "激活", new Date());
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

    // http://localhost:9022/test/import/simple
    @PostMapping("/import/simple")
    public List<User> importExcelSimple(@RequestParam("file") MultipartFile file) {
        try {
            InputStream is = file.getInputStream();
            List<User> users = ExcelUtil.importDataBrowser(is, User.class, true);
            users.forEach(item -> System.out.println(item));
            return users;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return new ArrayList<>();
    }

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
        return new ArrayList<>();
    }
}
