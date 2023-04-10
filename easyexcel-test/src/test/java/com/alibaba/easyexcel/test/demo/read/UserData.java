package com.alibaba.easyexcel.test.demo.read;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

/**
 * 基础数据类.这里的排序和excel里面的排序一致
 *  导入数据的类;
 * @author zhognweisheng
 **/
@Getter
@Setter
@EqualsAndHashCode

@ContentRowHeight(16)
@HeadRowHeight(43)
@ColumnWidth(12)
public class UserData {

    @ExcelProperty("* 登录名")
    private String username;
    @ExcelProperty("* 密码")
    private String password;
    @ExcelProperty("* 姓名")
    private String fullname;
    @ExcelProperty("* 英文名")
    private String realName;
    @ExcelProperty("手机号")
    private String phone;
    @ColumnWidth(50)
    @ExcelProperty("邮箱")
    private String email;
    @ExcelProperty("性别")
    private String gender;
    @ExcelProperty("部门")
    private String departmentName;
    @ExcelProperty("* 功能权限角色(多选)")
    private String roleName;
    @ExcelProperty("* 管理权限角色(多选)")
    private String manageName;
    @ExcelProperty("* 字段权限角色(多选)")
    private String fieldName;
    @ExcelProperty("* 是否接收询盘")
    private String canGetInquiry;//是否 1是2否
    @ExcelProperty("* 是否有网站聊天权限")
    private String canChat;//是否 1是2否
    @ExcelProperty("* 是否有WA企业账号聊天权限")
    private String canWabaService;//是否 1是2否
    @ExcelProperty("* 是否有Facebook聊天权限")
    private String canFbManage;//是否 1是2否
    @ExcelProperty("* 是否主管")
    private String isDeptManager;//是否 1是2否

}
