package online.dinghuiye.excel;


import online.dinghuiye.common.WriteMode;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import online.dinghuiye.api.AbstractExcel;

import java.io.File;
import java.io.InputStream;

/**
 * 创建.xlsx的{@link AbstractExcel}
 *
 * @author Strangeen
 * on 2017/7/5
 */
public class ExcelForXlsx extends AbstractExcel {

    public ExcelForXlsx(File excel) {
        super(excel);
    }

    public ExcelForXlsx(File excel, WriteMode mode) {
        super(excel, mode);
    }

    @Override
    protected void readWorkbook(InputStream fis) {
        try {
            this.wb = new XSSFWorkbook(fis);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    protected void createWorkbook() {
        try {
            this.wb = new XSSFWorkbook();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
