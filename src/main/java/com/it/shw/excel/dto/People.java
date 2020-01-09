package com.it.shw.excel.dto;

import com.it.shw.excel.annotion.ShowStyle;
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
