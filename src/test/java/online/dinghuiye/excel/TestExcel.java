package online.dinghuiye.excel;

import online.dinghuiye.api.AbstractExcel;
import online.dinghuiye.common.WriteMode;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.IntStream;

/**
 * 测试类
 *
 * @author Strangeen
 * on 2017/7/6
 */
public class TestExcel {

    @Test
    public void testExcelFactoryError() {

        File file = null;

        // Test null
        {
            try {
                ExcelFactory.newExcel(file);
            } catch (Exception e) {
                Assert.assertEquals("文件为空", e.getMessage());
            }
        }

        // Test not excel
        {
            try {
                file = new File("D:/test/test.txt");
                ExcelFactory.newExcel(file);
            } catch (Exception e) {
                Assert.assertEquals("文件不是excel", e.getMessage());
            }
        }
    }

    @Test
    public void testReadExcel() throws ParseException {
        AbstractExcel e = new ExcelForXls(new File("D:/test/test.xls"));
        testReadExcel(e);
    }

    @Test
    public void testReadExcelForXlsx() throws ParseException {
        AbstractExcel e = new ExcelForXlsx(new File("D:/test/test.xlsx"));
        testReadExcel(e);
    }

    private void testReadExcel(AbstractExcel e) throws ParseException {
        List<Map<String, Object>> res = e.readExcel(0);

        List<Object> actual = new ArrayList<>();
        res.forEach((Map<String, Object> map) -> actual.addAll(map.values()));

        List<Object> expect = Arrays.asList(new Object[]{
                new SimpleDateFormat("yyyy-MM-dd").parse("1970-1-2"),
                2.0, 123.0, 1.23, 11.0, 137.23000000000002, 666.0, "ab137.23"
        });

        IntStream.range(0,expect.size()).forEach(i -> Assert.assertEquals(expect.get(i), actual.get(i)));
    }



    @Test
    public void testCoverWriteExcel() throws InterruptedException {
        AbstractExcel e = new ExcelForXls(new File("D:/test/test_out.xls"));
        writeExcel(e);
    }

    @Test
    public void testCoverWriteExcelForXlsx() throws InterruptedException {
        AbstractExcel e = new ExcelForXlsx(new File("D:/test/test_out.xlsx"));
        writeExcel(e);
    }


    @Test
    public void testInsertWriteExcel() throws InterruptedException {
        AbstractExcel e = new ExcelForXls(new File("D:/test/test_out_insert.xls"), WriteMode.INSERT);
        writeExcel(e);
    }

    @Test
    public void testInsertWriteExcelForXlsx() throws InterruptedException {
        AbstractExcel e = new ExcelForXlsx(new File("D:/test/test_out_insert.xlsx"), WriteMode.INSERT);
        writeExcel(e);
    }


    private void writeExcel(AbstractExcel e) throws InterruptedException {
        // 测试数据
        List<List<String>> dataTDList = new ArrayList<>();
        List<String> titleNameList = new ArrayList<>();
        dataTDList.add(titleNameList);
        titleNameList.add("测试1");
        titleNameList.add("测试2");
        titleNameList.add("测试3");
        for (int i = 0; i < 10; i ++) {
            List<String> valueList = new ArrayList<>();
            dataTDList.add(valueList);
            if (i == 5) dataTDList.add(new ArrayList<>());
            for (int j = 0; j < titleNameList.size(); j ++) {
                valueList.add(String.valueOf(i + j));
            }
        }
        List<List<String>> dataTDList2 = new ArrayList<>();
        for (int i = 9; i < 20; i ++) {
            List<String> valueList = new ArrayList<>();
            dataTDList2.add(valueList);
            for (int j = 0; j < 5; j ++) {
                if (j == 3) {
                    valueList.add(null);
                    continue;
                }
                valueList.add(String.valueOf(i + j));
            }
        }

        e.writeExcel(dataTDList, "a_" + new Date().getTime(), false);
        e.writeExcel(dataTDList2, "", true);

        //Thread.sleep(30000); // 测试流是否被关闭
    }
}
