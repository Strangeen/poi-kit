package online.dinghuiye.excel;


import online.dinghuiye.common.WriteMode;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import online.dinghuiye.api.AbstractExcel;

import java.io.File;
import java.io.InputStream;

/**
 * Created by Strangeen on 2017/7/5.
 */
public class ExcelForXls extends AbstractExcel {

    public ExcelForXls(File excel) {
        super(excel);
    }

    public ExcelForXls(File excel, WriteMode mode) {
        super(excel, mode);
    }

    @Override
    protected void readWorkbook(InputStream fis) {
        try {
            this.wb = new HSSFWorkbook(fis);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    protected void createWorkbook() {
        try {
            this.wb = new HSSFWorkbook();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
