# **poi-kit**

apache poi封装，方便excel的读取和写入

## 获取jar包

maven下执行命令
```
mvn clean package -Dmaven.test.skip=true
```

## 1. 读取

### 描述

按行读取，每行按照`表头名称（首行内容）->各行对应内容`封装为`Map<String, Object>`，再将各行封装为`List`，最后得到`List<Map<String, Object>>`

其中，读取内容为`Object`，是根据excel单元格类型进行判：

1. 设置为`常规`，输入数字，读取为`long`或者`double`；输入的是文本，读取为`String`；输入的日期，读取为`Date`
2. 设置为`文本`，读取为`String`
3. 设置为`日期`，读取为`Date`

### 示例代码
```java
file = new File("D:/test.xls"); // excel文件位置，.xls或.xlsx
AbstractExcel excel = ExcelFactory.newExcel(file); // 创建AbstractExcel文件
List<Map<String, Object>> datas = excel.readExcel(0); // 读取sheet序号为0的数据
```

另外，还提供指定读取.xls或.xlsx的类
```java
AbstractExcel excel = new ExcelForXls(new File("D:/test.xls")); // .xls
```
```java
AbstractExcel excel = new ExcelForXlsx(new File("D:/test.xlsx")); // .xlsx
```

## 2. 写入

### 描述

将各列封装为`List`，再将各行封装为`List`，得到`List<List<String>>`，即为需要写入的数据集合，写入数据时，需要区分`插入sheet模式`和`覆盖sheet模式`

因为excel的写入并非主要目的，因此现在写入只支持`String`类型的写入，excel单元格将展示为`文本`格式，excel区分.xls和.xlsx，没有写工厂类

### 示例代码

```java
// 数据集合
List<List<String>> dataTDList = new ArrayList<>();
List<String> titleNameList = new ArrayList<>();
// 表头数据
dataTDList.add(titleNameList);
titleNameList.add("测试1");
titleNameList.add("测试2");
titleNameList.add("测试3");
// 内容数据
{
  List<String> valueList = new ArrayList<>();
  dataTDList.add(valueList);
  valueList.add(1);
  valueList.add(2);
  valueList.add(3);
}
{
  List<String> valueList = new ArrayList<>();
  dataTDList.add(valueList);
  valueList.add("a");
  valueList.add("1999-01-02");
  valueList.add("this is a test word");
}
// ...

// 创建AbstractExcel
AbstractExcel excel = new ExcelForXls(new File("D:/test.xls"), WriteMode.INSERT); // .xls 插入模式，即创建新的sheet
// AbstractExcel excel = new ExcelForXls(new File("D:/test.xls"), WriteMode.COVER); // 覆盖模式，覆盖已有的sheet，不传WriteMode默认为COVER
// AbstractExcel excel = new ExcelForXlsx(new File("D:/test.xlsx")); // .xlsx

// 写入数据
excel.writeExcel(dataTDList, "sheet name", true); // 写入数据
```

#### 关于`writeExcel`方法：

```java
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
  //...
}
```
