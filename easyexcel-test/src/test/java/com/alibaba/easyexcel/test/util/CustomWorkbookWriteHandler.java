//package com.alibaba.easyexcel.test.util;
//
//import com.alibaba.excel.write.handler.AbstractWorkbookWriteHandler;
//import com.alibaba.excel.write.handler.WorkbookWriteHandler;
//import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
//import org.apache.commons.collections4.CollectionUtils;
//import org.apache.commons.compress.utils.Lists;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//
//import java.util.List;
//
//public class CustomWorkbookWriteHandler extends WorkbookWriteHandler {
//
//    @Override
//    public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
//        Workbook workbook =  writeWorkbookHolder.getWorkbook();
//        int numberOfSheets = workbook.getNumberOfSheets();
//        List<String> removeList = Lists.newArrayList();
//        for (int i=0; i<numberOfSheets; i++) {
//            Sheet sheetAt = workbook.getSheetAt(i);
//            String sheetName = sheetAt.getSheetName();
//            if (sheetName.endsWith("删除")) {
//                removeList.add(sheetName);
//            }
//        }
//        if (CollectionUtils.isNotEmpty(removeList)) {
//            removeList.forEach(e -> {
//                int sheetIndex = workbook.getSheetIndex(e);
//                workbook.removeSheetAt(sheetIndex);
//            });
//        }
//    }
//}
