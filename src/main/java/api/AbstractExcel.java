package api;

import common.WriteMode;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Strangeen on 2017/7/6.
 *
 * poi版本3.16
 * 导入和导出分别创建AbstractExcel对象
 */
public abstract class AbstractExcel {

    protected File excel;
    protected InputStream fis;
    private OutputStream fos;
    protected List<String> titleNameList;

    protected Workbook wb;
    protected Sheet sheet;

    protected WriteMode mode;

    // 缓存Row
    private HashMap<Integer, Row> rowHashMap = new HashMap<Integer, Row>();


    /**
     * 导出模式为默认的覆盖 C
     */
    public AbstractExcel(File excel) {
        this.excel = excel;
        this.mode = WriteMode.C;
    }

    /**
     * 自定义导出模式
     * @param excel 导出的文件
     * @param mode 模式，如：创建sheet是覆盖文件uo
     *             C - 覆盖
     *             I - 插入
     */
    public AbstractExcel(File excel, WriteMode mode) {
        this.excel = excel;
        this.mode = mode;
    }

    public void setMode(WriteMode mode) {
        this.mode = mode;
    }

    public void setExcel(File excel) {
        this.excel = excel;
    }


    // ----------- 读取excel部分 ------------

    /**
     * 将excel的按行数据转换为Map<表头名，值>映射
     * @param sheetNo 读取sheet的编号
     * @return
     */
    public List<Map<String, String>> readExcel(int sheetNo) {

        try {
            List<Map<String, String>> recordMapList = new ArrayList<Map<String, String>>();
            open(excel, sheetNo);
            int totalRowNum = getTotalRowNum();
            for (int rowNo = 1; rowNo < totalRowNum; rowNo++) {
                Map<String, String> recordMap = new HashMap<String, String>();

                for (int colNo = 0; colNo < titleNameList.size(); colNo++) {
                    String value = readCell(rowNo, colNo);
                    recordMap.put(titleNameList.get(colNo), value);
                }
                recordMapList.add(recordMap);
            }

            return recordMapList;

        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            close();
        }
    }

    /**
     * 根据不同的excel类型来创建不同的wb
     * @param fis 文件流
     */
    protected abstract void readWorkbook(InputStream fis);

    /**
     * 打开excel文件，默认选中第一个sheet
     * @param excelFile xlsx文件
     */
    private void open(File excelFile, int sheetNo) {
        try {
            this.fis = new FileInputStream(excelFile);
            readWorkbook(fis);
            readSheet(sheetNo);
            readSheetTitleNameList();
        } catch (Exception e) {
            throw new RuntimeException(e);

        }
    }

    private void readSheet(int sheetNo) {
        try {
            this.sheet = wb.getSheetAt(sheetNo);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private int getTotalRowNum() {
        return sheet.getPhysicalNumberOfRows();
    }

    private void readSheetTitleNameList() {
        int firstRowNo = 0;
        Row firstRow = sheet.getRow(firstRowNo);
        this.titleNameList = new ArrayList<String>();
        if (firstRow != null) {
            for (int colNo = 0; colNo < firstRow.getPhysicalNumberOfCells(); colNo ++) {
                titleNameList.add(readCell(firstRowNo, colNo));
            }
        }
    }

    private String readCell(int rowNo, int colNo) {
        Row row = rowHashMap.get(rowNo);
        if (row == null) {
            row = sheet.getRow(rowNo);
            rowHashMap.put(rowNo, row);
        }
        Cell cell = row.getCell(colNo);
        if (cell == null) return null;
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue();
    }

    public void close() {
        try {
            if (wb != null && fos != null)
                // 因为代码结构问题，将wb写入excel在close里执行，
                // 否则writeExcel无法写入多张sheet
                wb.write(fos);

            if (wb != null) {
                wb.close();
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        try {
            if (fis != null)
                fis.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        try {
            if (fos != null)
                fos.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }



    // ----------- 写入excel部分 ------------

    /**
     * 将数据写到excel中
     * @param dataTDList 表格数据，按行列存入
     *                      如：第一个元素是表头名称List
     *                      后面的元素为数据List，顺序按表头名称顺序
     * @param sheetName sheet名称，null按默认名称Sheet0...
     * @param autoClose true 将wb写入excel文件，并自动关闭文件
     *                  false 手动关闭文件，注意：此时并没有将wb写入excel文件
     *                      可以手动调用close()方法将wb写入excel文件
     *                      创建多个sheet时需要保持文件不关闭
     *                      创建最后一个sheet时传入为true
     */
    public void writeExcel(List<List<String>> dataTDList, String sheetName, boolean autoClose) {
        try {
            if (wb == null) {
                if (mode == WriteMode.C) {
                    createWorkbook();
                } else if (mode == WriteMode.I) {
                    // 读取文件的wb作为要写入的wb
                    this.fis = new FileInputStream(excel);
                    readWorkbook(fis);
                } else {
                    throw new RuntimeException("导出模式设置错误");
                }
            }

            if (sheetName == null || sheetName.length() < 1 || sheetName.length() > 31)
                this.sheet = wb.createSheet();
            else {
                this.sheet = wb.createSheet(sheetName);
            }

            if (dataTDList.size() <= 0) {
                //throw new RuntimeException("没有数据，无法写入excel");
            }

            // 清空row缓存
            rowHashMap.clear();

            int rowNo = 0;
            for (List<String> dataList : dataTDList) {
                for (int colNo = 0; colNo < dataList.size(); colNo ++) {
                    String value = dataList.get(colNo);
                    writeCell(rowNo, colNo, value);
                }
                rowNo ++;
            }
            // 准备写入文件的流
            if (fos == null)
                this.fos = new FileOutputStream(excel);

            // 写入文件在close()中调用

        } catch (Exception e) {
            close();
            throw new RuntimeException(e);
        } finally {
            if (autoClose) {
                close();
            }
        }
    }

    protected abstract void createWorkbook();

    // 写入单元格
    private void writeCell(int rowNo, int colNo, String value) {
        Row row = rowHashMap.get(rowNo);
        if (row == null) {
            rowHashMap.put(rowNo, sheet.createRow(rowNo));
            row = rowHashMap.get(rowNo);
        }
        Cell cell = row.createCell(colNo);
        //cell.setCellType(CellType.STRING);
        cell.setCellValue(value);
    }
}
