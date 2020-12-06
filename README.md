# persistence-excel


Quick Start
-----------
## Maven/Gradle configuration

Add the Maven dependency:

```xml
<dependency>
    <groupId>com.happy3w</groupId>
    <artifactId>persistence-excel</artifactId>
    <version>0.0.3</version>
</dependency>
```

Add the Gradle dependency:

```groovy
implementation 'com.happy3w:persistence-excel:0.0.3'
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
MessageRecorder的用法可以参见 https://github.com/boroborome/toolkits#messagerecorder-%E7%BB%84%E4%BB%B6

ObjRdTableDef的使用方法参见 https://github.com/boroborome/persistence-core#objrdtabledef

#### 扩展数据类型转换规则
SheetPage默认使用系统toolkits中默认TypeConverter实例转换数据的副本，如果转换数据有自己的要求，可以直接设置新的typeConverter。

#### 扩展不同数据类型的读写方式
SheetPage.cellAccessManager负责配置如何将不同数据类型读写到一个单元格中，

系统默认支持了String、Long、Integer、Date，如果出现其他自己的类型，可以扩展这个cellAccessManager

默认情况SheetPage.cellAccessManager会从CellAccessManager.INSTANCE复制一份，因此可以配置全局的CellAccessManager.INSTANCE来让有所的SheetPage生效。也可以单独修改一个SheetPage实例的cellAccessManager。举例说明如何实现新的ICellAccessor

```java
// 这是负责读写Double类型的Accessor
public class DoubleCellAccessor implements ICellAccessor<Double> {
    // 将值写入一个单元格
    @Override
    public void write(Cell cell, Double value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    // 从给定的cell读取一个Double类型数据
    @Override
    public Double read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
            return null;
        }
        return cell.getNumericCellValue();
    }

    @Override
    public Class<Double> getType() {
        return Double.class;
    }
}

// 在适当的实际将这个CellAccessor添加到管理器
CellAccessManager.INSTANCE.registItem(new DoubleCellAccessor());

```

#### 扩展Excel支持的IRdConfig
SheetPage支持的IRdConfig配置类型，可以通过SheetPage.regRdConfigInfo方法增加或者修改。

如果需要修改默认值，可以直接修改RdciHolder.ALL_CONFIG_INFOS。

一个RdConfigInfo是一个让SheetPage认识IRdConfig的配置信息，详细描述如下：

```java
/**
 * 配置处理信息。SheetPage用这些信息让各种配置在Excel上生效
 * @param <VT> 需要处理数据的数据类型
 * @param <CT> 对应配置的类型
 */
@Getter
public abstract class RdConfigInfo<VT, CT extends IRdConfig> implements ITypeItem<VT> {
    /**
     * 需要处理数据的数据类型。如果对数据没有要求就填写Void.class。在同一个单元格上非Void.class的配置信息之间会只有一个生效。系统会根据优先级在其中选择一个。
     */
    protected Class<VT> type;

    /**
     * 对应配置的类型。
     */
    protected Class<CT> configType;

    public RdConfigInfo(Class<CT> configType) {
        this(configType, null);
    }

    public RdConfigInfo(Class<CT> configType, Class<VT> type) {
        this.type = type;
        this.configType = configType;
        this.isDataFormat = type != null && type != Void.class && type != Object.class;
    }

    /**
     * 标示当前配置是否属于格式配置<br>
     *     一个Cell上只能有一个格式配置，因此，多个格式配置之间互相冲突。使用的时候只有优先级最高的生效<br>
     *     生效顺序：Cell上配置、Column配置、Sheet上配置。分别对应某次写入是特别指定的配置，字段上的注解，类上的注解与Page上的配置。
     */
    protected boolean isDataFormat;

    /**
     * 将这个配置应用到CellStyle上
     * @param cellStyle 等待配置的cellStyle
     * @param rdConfig 需要配置到cellStyle上的配置信息
     * @param bsc 包含当前单元格信息的一些上下文
     */
    public abstract void buildStyle(CellStyle cellStyle, CT rdConfig, BuildStyleContext bsc);
}
```
当前只有DateFormatCfg和NumFormatCfg的配置信息属于数据格式类配置，其他都是普通的样式配置。

#### 扩展Excel特有的配置
当前工具只增加了下面注解
- FillForegroundColor
  背景色配置。支持的颜色定义在HssfColor
- FillPattern
  背景填充方式配置。一般使用FillPatternType.SOLID_FOREGROUND。

如果需要扩展，可以自行实现。
- 在代码中实现自己的注解和配置，参见 https://github.com/boroborome/persistence-core#%E6%89%A9%E5%B1%95%E6%B3%A8%E8%A7%A3
- 实现新注解的RdConfigInfo。参见上面《扩展Excel支持的IRdConfig》

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
## 历史
### 0.0.4
- 升级libs，修复数据转换的bug
