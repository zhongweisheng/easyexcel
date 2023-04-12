package com.alibaba.easyexcel.test.util;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CustomSheetWriteHandlerNew implements SheetWriteHandler {


    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        //定义一个map key是需要添加下拉框的列的index value是下拉框数据
        Map<Integer, String[]> mapDropDown = new HashMap<>(4);
        //设置单位身份 值写死
        String[] unitIdentity = {"职教集团成员单位", "拟合作单位", "合作单位"};

        //地区
        String[] area = {"国内", "国外"};

        //等级下拉选
        String[] level = {"国家级示范校", "国家级骨干校", "省级示范校", "省级骨干校", "其他"};

        //年份下拉选
//        String[] joinYear = {"1990年","1991年","1992年","1993年","1994年","1995年","1996年","1997年","1998年",
//                "1999年","2000年","2001年","2002年","2003年","2004年","2005年","2006年","2007年","2008年","2009年","2010年","2011年","2012年",
//                "2013年","2014年","2015年","2016年","2017年","2018年","2019年","2020年","2021年","2022年"};

        String[] joinYear = {
            "10171主管",
            "10171sale",
            "10171销售",
            "10171测试1",
            "10171客服",
            "全菜单10171",
            "10171仅聊天",
            "默认管理员",
            "销售",
            "客服",
            "销售经理",
            "市场经理",
            "销售人员",
            "我的测试",
            "fb权限申请请勿动",
            "yd测试专用0626",
            "去掉facebook",
            "无线索",
            "王芳测试勿动",
            "测试二维码",
            "aaaa",
            "测试A",
            "无公海权限",
            "没有话术设置菜单",
            "没有WA功能权限",
            "无公海管理_线索功能权限",
            "无公海管理_线索功能权限",
            "无公海管理_线索功能权限",
            "王志民测试公海规则角色001",
            "王志民测试公海规则角色002",
            "王志民测试公海规则角色003",
            "王志民测试公海规则角色004",
            "王志民测试公海规则角色005"};

        //下拉选在Excel中对应的列
        mapDropDown.put(0, unitIdentity);
        mapDropDown.put(1, joinYear);
        mapDropDown.put(4, area);
        mapDropDown.put(10, level);

        //获取工作簿
        Sheet sheet = writeSheetHolder.getSheet();
        ///开始设置下拉框
        DataValidationHelper helper = sheet.getDataValidationHelper();

//获取一个workbook
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        //定义sheet的名称
        String hiddenName = "hidden" + System.currentTimeMillis();

        //1.创建一个隐藏的sheet 名称为 hidden
        Sheet hidden = workbook.createSheet(hiddenName);
        workbook.setSheetHidden(workbook.getSheetIndex(hidden),true );

        //设置下拉框
        for (Map.Entry<Integer, String[]> entry : mapDropDown.entrySet()) {
            String[] valueArr = entry.getValue();
            Integer key = entry.getKey();
            if (valueArr.length < 20) {
                /*起始行、终止行、起始列、终止列  起始行为1即表示表头不设置**/
                CellRangeAddressList addressList = new CellRangeAddressList(1, 65535, entry.getKey(), entry.getKey());
                /*设置下拉框数据**/
                DataValidationConstraint constraint = helper.createExplicitListConstraint(entry.getValue());
                DataValidation dataValidation = helper.createValidation(constraint, addressList);
                sheet.addValidationData(dataValidation);
            } else {
                int rowNo = 100;
                //2.循环赋值（为了防止下拉框的行数与隐藏域的行数相对应，将隐藏域加到结束行之后）
                for (int i = 0, length = valueArr.length; i < length; i++) {
                    // 3:表示你开始的行数  3表示 你开始的列数
                    hidden.createRow(rowNo + i).createCell(key).setCellValue(valueArr[i]);
                }
                Name category1Name = workbook.createName();
                category1Name.setNameName(hiddenName);
                //4 A1:A代表隐藏域创建第N列createCell(N)时。以A1列开始A行数据获取下拉数组
                category1Name.setRefersToFormula(hiddenName + "!A1:A" + (valueArr.length + rowNo));
                //5 将刚才设置的sheet引用到你的下拉列表中
                DataValidationConstraint constraint8 = helper.createFormulaListConstraint(hiddenName);
                //5 将刚才设置的sheet引用到你的下拉列表中
                CellRangeAddressList addressList = new CellRangeAddressList(1, 65535, entry.getKey(), entry.getKey());
                DataValidation dataValidation3 = helper.createValidation(constraint8, addressList);
                sheet.addValidationData(dataValidation3);
            }
        }


    }

}

