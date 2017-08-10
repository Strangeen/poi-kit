package online.dinghuiye.excel;

import online.dinghuiye.api.AbstractExcel;

import java.io.File;

/**
 * 通过Factory创建{@link AbstractExcel}，透明创建.xls和.xlsx为后缀的{@link AbstractExcel}
 *
 * @author Strangeen
 * on 2017/08/08
 */
public class ExcelFactory {

    public static AbstractExcel newExcel(File file) {

        if (file == null) throw new RuntimeException("文件为空");
        String fileName = file.getName();
        if (fileName.endsWith(".xls"))
            return new ExcelForXls(file);
        else if (fileName.endsWith(".xlsx"))
            return new ExcelForXlsx(file);
        else
            throw new RuntimeException("文件不是excel");
    }

}
