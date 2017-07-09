package excel;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import api.AbstractExcel;

import java.io.File;
import java.io.InputStream;

/**
 * Created by Strangeen on 2017/7/5.
 */
public class ExcelForXls extends AbstractExcel {

    /**
     * 创建Excel，读写分开创建对象
     * @param xls 读取或保存到的文件
     */
    public ExcelForXls(File xls) {
        super.excel = xls;
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
