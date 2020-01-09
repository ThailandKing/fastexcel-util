package com.it.shw.excel.dto;

import com.it.shw.excel.annotion.ShowStyle;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * @Copyright: Harbin Institute of Technology.All rights reserved.
 * @Description:
 * @author: thailandking
 * @since: 2020/1/7 14:46
 * @history: 1.2020/1/7 created by thailandking
 */
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
