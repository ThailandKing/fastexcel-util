package com.it.shw.excel;

import com.it.shw.excel.dto.User;
import com.it.shw.excel.utils.ExcelUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class SpringbootExcelDemoApplicationTests {

    @Test
    public void test() {
        String str = "1";
        Long l = Long.valueOf(str);
        System.out.println(l);
        System.out.println(new Date());
    }
}
