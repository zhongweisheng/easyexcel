package com.alibaba.easyexcel.test.demo.write;

import com.alibaba.easyexcel.test.demo.read.UserData;
import com.alibaba.easyexcel.test.util.CustomSheetWriteHandlerNew;
import com.alibaba.easyexcel.test.util.TestFileUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.Test;

import java.util.List;

/**
 * @PROJECT_NAME: easyexcel-parent
 * @DESCRIPTION:
 * @USER: zhongweisheng
 * @DATE: 2023/3/9 10:10
 */
public class UserWriterTest {

    @Test
    public void simpleWrite() {
        // 注意 simpleWrite在数据量不大的情况下可以使用（5000以内，具体也要看实际情况），数据量大参照 重复多次写入

        // 写法1 JDK8+
        // since: 3.0.0-beta1
        String fileName = TestFileUtil.getPath() + "UserDatasimpleWrite" + System.currentTimeMillis() + ".xlsx";

        System.out.println(  fileName  );
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        // 如果这里想使用03 则 传入excelType参数即可
//        EasyExcel.write(fileName, UserData.class)
//            .sheet("模板")
//            .doWrite(() -> {
//                // 分页查询数据
//                return data();
//            });

        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为红色

//     final String DEFAULT_BACKGROUND_COLOR = "#BDD8EE";

        // 自定义背景色
//        int r = Integer.parseInt((DEFAULT_BACKGROUND_COLOR.substring(1,3)),16);
//        int g = Integer.parseInt((DEFAULT_BACKGROUND_COLOR.substring(3,5)),16);
//        int b = Integer.parseInt((DEFAULT_BACKGROUND_COLOR.substring(5,7)),16);
//
//        HSSFWorkbook wb = new HSSFWorkbook();
//        HSSFPalette palette = wb.getCustomPalette();
//        HSSFColor hssfColor = palette.findSimilarColor(r, g, b);
//        headWriteCellStyle.setFillForegroundColor(hssfColor.getIndex());

        headWriteCellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        headWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 11);
        headWriteCellStyle.setWriteFont(headWriteFont);
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short) 10);
        contentWriteFont.setFontName("等线");
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
            new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);

        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        EasyExcel.write(fileName, UserData.class)
//            .registerWriteHandler(horizontalCellStyleStrategy)
//            .registerWriteHandler(new CustomSheetWriteHandlerNew())
//            .sheet("模板")
//            .doWrite(data());

        try (ExcelWriter excelWriter = EasyExcel.write(fileName, UserData.class)
            .registerWriteHandler(horizontalCellStyleStrategy)
            .registerWriteHandler(new CustomSheetWriteHandlerNew()).build()) {
            // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
            for (int i = 0; i < 5; i++) {
                // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样
                WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).build();
                // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
                List<UserData> data = data();
                excelWriter.write(data, writeSheet);
            }
        }


//        ExcelWriter excelWriter = EasyExcel.write(fileName, UserData.class)
//            .registerWriteHandler(horizontalCellStyleStrategy)
//            .registerWriteHandler(new CustomSheetWriteHandlerNew())
//            .build();
//
////            WriteSheet writeSheet = EasyExcel.writerSheet(0, "模板" + 0).build();
//            WriteSheet writeSheet = EasyExcel.writerSheet(0, "模板" + 0).build();
//            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
//            List<UserData> data = data();
//            excelWriter.write(data, writeSheet);

//        // 方法2: 如果写到不同的sheet 同一个对象
//        fileName = TestFileUtil.getPath() + "repeatedWrite" + System.currentTimeMillis() + ".xlsx";
//        // 这里 指定文件
//        try (ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build()) {
//            // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
//            for (int i = 0; i < 5; i++) {
//                // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样
//                WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).build();
//                // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
//                List<DemoData> data = data();
//                excelWriter.write(data, writeSheet);
//            }
//        }


//        WriteSheet writeSheet1 = EasyExcel.writerSheet(1, "模板" + 1).build();
//        excelWriter.write(data, writeSheet1);

//        // 写法2
//        fileName = TestFileUtil.getPath() + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        // 如果这里想使用03 则 传入excelType参数即可
//        EasyExcel.write(fileName, UserData.class).sheet("模板").doWrite(data());

//        // 写法3
//        fileName = TestFileUtil.getPath() + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
//        // 这里 需要指定写用哪个class去写
//        try (ExcelWriter excelWriter = EasyExcel.write(fileName, UserData.class).build()) {
//            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
//            excelWriter.write(data(), writeSheet);
//        }
    }

    private List<UserData> data() {
        List<UserData> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            UserData user = new UserData();
            user.setCanChat("是");
            user.setCanFbManage("是");
            user.setCanGetInquiry("是");
            user.setCanWabaService("是");
            user.setDepartmentName("销售部");
            user.setEmail("xunpanyun@xhlmarketing.com");
            user.setFieldName("默认管理员");
            user.setFullname("询盘云");
            user.setGender("男");
            user.setIsDeptManager("是");
            user.setManageName("本公司");
            user.setPassword("xunpanyun188");
            user.setPhone("18888888888");
            user.setRealName("leadscloud");
            user.setRoleName("默认管理员");
            user.setUsername("xunpanyun");
            list.add(user);
        }
        return list;
    }


}
