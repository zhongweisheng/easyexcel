package com.alibaba.easyexcel.test.demo.zws;

import com.alibaba.easyexcel.test.util.TestFileUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.junit.Test;

import java.util.ArrayList;

/**
 * @PROJECT_NAME: easyexcel-parent
 * @DESCRIPTION:
 * @USER: zhongweisheng
 * @DATE: 2023/4/10 15:53
 */
public class MyZwsTest {

    @Test
    public void simpleWrite() {
        String fileName = "e:/test-"+System.currentTimeMillis()+".xlsx";

        ExcelWriter excelWriter = EasyExcel.write(fileName).build();

        WriteSheet writeSheet = EasyExcelUtil.writeSelectedSheet(UserDTO.class, 0, "测试sheet");
        excelWriter.write(new ArrayList<String>(), writeSheet);
        excelWriter.finish();
    }
}
