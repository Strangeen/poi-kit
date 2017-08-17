package online.dinghuiye.api;

import online.dinghuiye.common.WriteMode;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * poi版本3.16
 * 导入和导出分别创建AbstractExcel对象
 *
 * @apiNote
 *  <p>poi版本3.16</p>
 *  <p>excel单元格类型按照{@link CellType}进行判断，{@link CellType#FORMULA}格式获取为公式计算后的值</p>
 *
 * @author Strangeen
 * on 2017/7/6
 */
public abstract class AbstractExcel {

    protected File excel;
    private InputStream fis;
    private OutputStream fos;
    private List<String> titleNameList;

    protected Workbook wb;
    private Sheet sheet;

    private WriteMode mode;

    private FormulaEvaluator evaluator;

    // 缓存Row
    private HashMap<Integer, Row> rowHashMap = new HashMap<>();


    /**
     * 导出模式为默认的覆盖 COVER
     */
    public AbstractExcel(File excel) {
        this.excel = excel;
        this.mode = WriteMode.COVER;
    }

    /**
     * 自定义导出模式
     *
     * @param excel 导出的文件
     * @param mode 模式，如：创建sheet是覆盖文件还是在同文件中插入sheet
     *             COVER - 覆盖
     *             INSERT - 插入
     */
    public AbstractExcel(File excel, WriteMode mode) {
        this.excel = excel;
        this.mode = mode;
    }

//    public void setMode(WriteMode mode) {
//        this.mode = mode;
//    }

    public void setExcel(File excel) {
        this.excel = excel;
    }


    // ----------- 读取excel部分 ------------

    /*
     * 将excel的按行数据转换为Map<表头名, 值>映射
     * @param sheetNo 读取sheet的编号
     * @return Map<表头名, 值>的List
     * @deprecated 读取单元格方法已替换为
     *               {@link AbstractExcel#readCellToObj(int, int)}
     */
    /*public List<Map<String, String>> readExcel(int sheetNo) {

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
    }*/


    /**
     * 读取excel
     *
     * @param sheetNo sheet编号
     * @return 每行数据的 表头名称，值 映射
     */
    public List<Map<String, Object>> readExcel(int sheetNo) {

        try {
            List<Map<String, Object>> recordMapList = new ArrayList<>();
            open(excel, sheetNo);
            int totalRowNum = getTotalRowNum();
            for (int rowNo = 1; rowNo < totalRowNum; rowNo++) {
                Map<String, Object> recordMap = new HashMap<>();

                for (int colNo = 0; colNo < titleNameList.size(); colNo++) {
                    Object value = readCellToObj(rowNo, colNo);
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
     *
     * @param fis 文件流
     */
    protected abstract void readWorkbook(InputStream fis);

    /**
     * 打开excel文件，默认选中第一个sheet
     *
     * @param excelFile xlsx文件
     */
    private void open(File excelFile, int sheetNo) {
        try {
            this.fis = new FileInputStream(excelFile);
            readWorkbook(fis);
            this.evaluator = wb.getCreationHelper().createFormulaEvaluator(); // 设置公式计算器
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
        this.titleNameList = new ArrayList<>();
        if (firstRow != null) {
            for (int colNo = 0; colNo < firstRow.getPhysicalNumberOfCells(); colNo ++) {
                titleNameList.add(readCellToString(firstRowNo, colNo));
            }
        }
    }

    /**
     * 读取sheet表头单元格，值类型均为字符串类型
     *
     * @param rowNo 行号
     * @param colNo 列号
     * @return 单元格值
     */
    private String readCellToString(int rowNo, int colNo) {
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

    /**
     * <p>读取数据单元格，值类型为对应的对象类型，
     * 公式类型{@link CellType#FORMULA}单元格的值为公式计算结果值</p>
     *
     * @param rowNo 行号
     * @param colNo 列号
     * @return 单元格值对应单元格类型的对象
     */
    private Object readCellToObj(int rowNo, int colNo) {
        try {
            Row row = rowHashMap.get(rowNo);
            if (row == null) {
                row = sheet.getRow(rowNo);
                rowHashMap.put(rowNo, row);
            }
            if (row == null) return null; // 从sheet中重新获取依然为null，那么返回null
            Cell cell = row.getCell(colNo);
            if (cell == null) return null;
            return getCellValueAccordCellType(cell);
        } catch (Exception e) {
            throw new RuntimeException("类型错误, 行：" + (rowNo + 1) + "，列：" + (colNo + 1), e);
        }
    }

    /**
     * <p>根据{@link CellType}获取单元格的值</p>
     * <p>特殊的公式格式，先转换为计算结果的类型，再递归获得计算值</p>
     *
     * @param cell 单元格
     * @return 单元格值对应单元格类型的对象
     */
    private Object getCellValueAccordCellType(Cell cell) {
        if (cell == null) return null;
        CellValue cellValueCell = evaluator.evaluate(cell);
        if (cellValueCell == null) return null;
        Object cellValue;
        switch (cellValueCell.getCellTypeEnum()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = cell.getDateCellValue();
                }
                else {
                    cellValue = cell.getNumericCellValue();
                    // 判断long还是double
                    cell.setCellType(CellType.STRING);
                    String cellStrValue = cell.getStringCellValue();
                    if (cellStrValue != null && !cellStrValue.contains(".")) // Long
                        cellValue = ((Double) cellValue).longValue();
                }
                break;
            case BLANK:
                cellValue = "";
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                throw new RuntimeException("formula is impossible");
                //cellValue = getCellValueAccordCellType(cell, evaluator.evaluate(cell).getCellTypeEnum());
            case ERROR:
                throw new RuntimeException("类型错误");
            default:
                cell.setCellType(CellType.STRING);
                cellValue = cell.getStringCellValue();
        }
        return cellValue;
    }


    private void close() {
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
     *
     * @param dataTDList 表格数据，按行列存入
     *                      如：第一个元素是表头名称List
     *                      后面的元素为数据List，顺序按表头名称顺序
     * @param sheetName sheet名称，null或空串或字数小于1大于31，均按默认名称Sheet0...
     * @param autoClose true 将wb写入excel文件，并自动关闭文件
     *                  false 手动关闭文件，注意：此时并没有将wb写入excel文件
     *                      可以手动调用close()方法将wb写入excel文件
     *                      创建多个sheet时需要保持文件不关闭
     *                      创建最后一个sheet时传入为true
     */
    public void writeExcel(List<List<String>> dataTDList, String sheetName, boolean autoClose) {
        try {
            if (wb == null) {
                if (mode == WriteMode.COVER) {
                    createWorkbook();
                } else if (mode == WriteMode.INSERT) {
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

//            if (dataTDList.size() <= 0) {
//                throw new RuntimeException("没有数据，无法写入excel");
//            }

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
