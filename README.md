# persistence-excel


Quick Start
-----------
## Maven/Gradle configuration

Add the Maven dependency:

```xml
<dependency>
    <groupId>com.happy3w</groupId>
    <artifactId>persistence-excel</artifactId>
    <version>0.0.1</version>
</dependency>
```

Add the Gradle dependency:

```groovy
implementation 'com.happy3w:persistence-excel:0.0.1'
```

## 组件介绍
- SheetPage Excel Sheet页
- ExcelAssistant Excel功能助手

---

### SheetPage
这里是一个Demo，先定义自己的数据结构
```java
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public static class MyData {
    @ObjRdColumn("名字")      // 声明Excel中的标题
    private String name;

    @ObjRdColumn("年龄")
    private int age;
}
```

写Excel的逻辑
```java
List<MyData> dataList = getDatas(); //拿到需要操作的数据

// 创建Excel workbook，以及用于保存数据的sheet页
Workbook workbook = ExcelUtil.newXlsWorkbook();
SheetPage page = SheetPage.of(workbook, "test-page");

// 生成数据定义，并将数据写入到page中
ObjRdTableDef<MyData> tableDef = ObjRdTableDef.from(MyData.class);
RdAssistant.writeObj(dataList.stream(), page, tableDef);

// 将Excel写入到文件
workbook.write(new FileOutputStream(excelFile));
```

读Excel数据
```java
// 打开excel文件，并获取包含数据的sheet页test-page
Workbook workbook = ExcelUtil.openWorkbook(new FileInputStream(excelFile));
SheetPage page = SheetPage.of(workbook, "test-page");

// 读取所有数据
MessageRecorder messageRecorder = new MessageRecorder();
List<MyData> datas = RdAssistant.readObjs(objRdTableDef, page, messageRecorder)
        .collect(Collectors.toList());

messageRecorder.getErrors(); // 所有解析文件过程中的错误信息
messageRecorder.isSuccess(); // 检测解析过程是否成功
```
MessageRecorder的用法可以参见https://github.com/boroborome/toolkits

### ExcelAssistant
通过ExcelAssistant可以读取所有Sheet页内容
```java
// 打开excel文件，
Workbook workbook = ExcelUtil.openWorkbook(new FileInputStream(excelFile));

// 读取所有Sheet页中数据。要求每页都有标题，各页的标题顺序可以不同
MessageRecorder messageRecorder = new MessageRecorder();
List<MyData> datas = ExcelAssistant.readRows(objRdTableDef, workbook, messageRecorder)
        .collect(Collectors.toList());

```
