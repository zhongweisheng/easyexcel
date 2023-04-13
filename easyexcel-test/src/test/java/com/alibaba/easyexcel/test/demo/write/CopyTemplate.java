//package com.alibaba.easyexcel.test.demo.write;
//
//import com.alibaba.excel.EasyExcel;
//import com.alibaba.excel.ExcelWriter;
//import com.alibaba.excel.write.metadata.WriteSheet;
//import com.alibaba.excel.write.metadata.fill.FillConfig;
//import com.alibaba.excel.write.metadata.fill.FillWrapper;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import javax.servlet.http.HttpServletResponse;
//import java.io.*;
//import java.util.ArrayList;
//import java.util.HashMap;
//import java.util.List;
//import java.util.Map;
//
///**
// * @PROJECT_NAME: easyexcel-parent
// * @DESCRIPTION:
// * @USER: zhongweisheng
// * @DATE: 2023/3/9 16:20
// */
//public class CopyTemplate {
//
//    public void fillTemplate(String templateFileName, List list, HttpServletResponse response, ExcelWriter excelWriter) {
//        //excel模板
//        File file = new File(templateFileName);
//        try (FileInputStream fileInputStream = new FileInputStream(file); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
//
//            //通过poi复制出需要的sheet个数的模板
//            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
//            workbook.setSheetName(0, "统计表");
//
//            //通过循环复制模板生成sheet，从1开始是因为我要把 list.get(0)在后面填充到sheet1中
//            for (int i = 1; i < list.size(); i++) {
//                Map<String, Object> map = (Map<String, Object>) list.get(i);
//                //复制sheet1模板，得到第i个sheet
//                workbook.cloneSheet(1, map.get("name")+""+i);
//            }
//
//            Map<String, Object> map2 = (Map<String, Object>) list.get(0);
//            String name = map2.get("name") + "" + 0;
//            //给模板设置sheet1的名称为name
//            workbook.setSheetName(1,name);
//
//            //写到流里
//            workbook.write(bos);
//            byte[] bArray = bos.toByteArray();
//            InputStream is = new ByteArrayInputStream(bArray);
//            //输出文件路径
//            excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(is).build();
//
//            //sheet0的合并单元格策略map
////            Map<String, List<RowRangeDto>> strategyMap = ExcelUtil.addMerStrategy(list);
////            WriteSheet writeSheet = EasyExcel.writerSheet("统计表")
////                .registerWriteHandler(new BizMergeStrategy(strategyMap)).build();
//            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
////            //sheet0填充数据
////            excelWriter.fill(new FillWrapper("list", list), fillConfig, writeSheet);
//
//            //sheet0后面的sheet循环填充数据
//            for (int i = 0; i < list.size(); i++) {
//                Map<String, Object> map = (Map<String, Object>) list.get(i);
//
//                //MyHandler，填充时延续列表的单元格格式，否则只有第一列会延续模板的格式，具体方法，贴在后面链接
//                WriteSheet writeSheet2 = EasyExcel.writerSheet(map.get("name")+""+i).registerWriteHandler(new MyHandler()).build();
//                Map<String, Object> singleMap = new HashMap<>();
//
//                singleMap.put("deptName",StringUtil.isNotEmpty(map.get("deptName"))?map.get("deptName"):"");
//
//                excelWriter.fill(singleMap, writeSheet2);
//                ArrayList<Map<String, Object>> mingpaiList = (ArrayList<Map<String, Object>>) map.get("mingpaiList");
//                excelWriter.fill(new FillWrapper("mingpailist", mingpaiList), fillConfig, writeSheet2);
//            }
//            // 关闭流
//            excelWriter.finish();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//
//
//}
