package online.dinghuiye.excel;

import online.dinghuiye.common.WriteMode;
import org.junit.Test;
import online.dinghuiye.api.AbstractExcel;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by Strangeen on 2017/7/6.
 */
public class TestExcel {

    @Test
    public void testReadExcel() {
        AbstractExcel e = new ExcelForXls(new File("D:/test/test.xls"));
        System.out.println(e.readExcel(0));
    }

    @Test
    public void testReadExcelForXlsx() {
        AbstractExcel e = new ExcelForXlsx(new File("D:/test/test.xlsx"));
        System.out.println(e.readExcel(0));
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
        AbstractExcel e = new ExcelForXls(new File("D:/test/test_out_insert.xls"), WriteMode.I);
        writeExcel(e);
    }

    @Test
    public void testInsertWriteExcelForXlsx() throws InterruptedException {
        AbstractExcel e = new ExcelForXlsx(new File("D:/test/test_out_insert.xlsx"), WriteMode.I);
        writeExcel(e);
    }


    private void writeExcel(AbstractExcel e) throws InterruptedException {
        // 测试数据
        List<List<String>> dataTDList = new ArrayList<List<String>>();
        List<String> titleNameList = new ArrayList<String>();
        dataTDList.add(titleNameList);
        titleNameList.add("测试1");
        titleNameList.add("测试2");
        titleNameList.add("测试3");
        for (int i = 0; i < 10; i ++) {
            List<String> valueList = new ArrayList<String>();
            dataTDList.add(valueList);
            if (i == 5) dataTDList.add(new ArrayList<String>());
            for (int j = 0; j < titleNameList.size(); j ++) {
                valueList.add(String.valueOf(i + j));
            }
        }
        List<List<String>> dataTDList2 = new ArrayList<List<String>>();
        for (int i = 9; i < 20; i ++) {
            List<String> valueList = new ArrayList<String>();
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
