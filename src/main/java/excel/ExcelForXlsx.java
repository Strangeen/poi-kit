package excel;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import api.AbstractExcel;

import java.io.File;
import java.io.InputStream;

/**
 * Created by Strangeen on 2017/7/5.
 */
public class ExcelForXlsx extends AbstractExcel {

    public ExcelForXlsx(File xlsx) {
        super.excel = xlsx;
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
