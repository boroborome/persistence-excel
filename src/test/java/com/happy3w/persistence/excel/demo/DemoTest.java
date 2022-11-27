package com.happy3w.persistence.excel.demo;

import com.alibaba.fastjson.JSON;
import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.persistence.core.rowdata.RdAssistant;
import com.happy3w.persistence.core.rowdata.config.DateFormatCfg;
import com.happy3w.persistence.core.rowdata.config.DateZoneIdCfg;
import com.happy3w.persistence.core.rowdata.config.NumFormatCfg;
import com.happy3w.persistence.core.rowdata.obj.ObjRdTableDef;
import com.happy3w.persistence.core.rowdata.simple.ListRdTableDef;
import com.happy3w.persistence.core.rowdata.simple.RdColumnDef;
import com.happy3w.persistence.excel.ExcelUtil;
import com.happy3w.persistence.excel.SheetPage;
import com.happy3w.persistence.excel.rdci.FillForegroundColorCfg;
import com.happy3w.persistence.excel.util.HssfColor;
import com.happy3w.toolkits.message.MessageRecorder;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class DemoTest {
    @Test
    public void should_read_write_by_ann_success() throws IOException {
        List<Student> orgStudentList = createTestData();

        // ----------写数据方法
        // orgStudentList是保存Student数据的列表

        // 创建一个Excel workbook，并创建一个test-page Sheet页面用于保存数据
        Workbook workbook = ExcelUtil.newXlsxWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");

        // 重点：通过Student创建一个"行数据表定义"
        ObjRdTableDef<Student> objRdTableDef = ObjRdTableDef.from(Student.class);

        // 通过"行数据助理"将数据写入excel page
        RdAssistant.writeObj(page, orgStudentList.stream(), objRdTableDef);

        workbook.write(Files.newOutputStream(new File("test.xlsx").toPath()));

        // ---------读数据
        // 从文件或者什么流中读入workbook。这里同时支持xlsx或者xls格式
        Workbook readWorkbook = ExcelUtil.openWorkbook(new FileInputStream(new File("test.xlsx")));
        SheetPage readPage = SheetPage.of(readWorkbook, "test-page");

        // 创建用于接收错误信息的recorder
        MessageRecorder messageRecorder = new MessageRecorder();

        // 将page中所有数据读取出来，错误信息记录到recorder中
        List<Student> newDataList = RdAssistant.readObjs(readPage, objRdTableDef, messageRecorder)
                .collect(Collectors.toList());

        if (messageRecorder.isSuccess()) {
            // 保存数据到数据库
        } else {
            // messageRecorder.getErrors();	// 读取所有错误信息，如果Excel中有多个错误，这里是多个错误
            // messageRecorder.toResponse();// 将errors,warnings等各种信息转换为一个response返回
        }

        Assertions.assertEquals(JSON.toJSONString(orgStudentList),
                JSON.toJSONString(newDataList));
    }

    private List<Student> createTestData() {
        List<Student> orgStudentList = Arrays.asList(
                Student.builder().name("Tom")
                        .age(12)
                        .studying(true)
                        .weight(55.4)
                        .updateTime(Timestamp.valueOf("2020-10-10 23:00:00").getTime())
                        .build(),
                Student.builder().name("张三")
                        .birthday(Timestamp.valueOf("2020-10-10 23:00:00"))
                        .build());
        return orgStudentList;
    }

    private ExtConfigs createExtConfigs(IRdConfig... configs) {
        ExtConfigs extConfigs = new ExtConfigs();
        for (IRdConfig config : configs) {
            extConfigs.regist(config);
        }
        return extConfigs;
    }

    @Test
    public void should_read_write_success_with_rddef() throws IOException {
        List<Student> orgStudentList = createTestData();

        ListRdTableDef rdTableDef = new ListRdTableDef();
        rdTableDef.config(new NumFormatCfg("#.00"))
                .setColumns(Arrays.asList(RdColumnDef.builder() // 按照Excel中出现的Title顺序填写
                                .title("名字")
                                .dataType(String.class)         // 数据类型是用于读取Excle用的，如果只用于写入，可以不填写这个信息
                                .extConfigs(createExtConfigs(new FillForegroundColorCfg(HssfColor.RED)))
                                .build(),
                        RdColumnDef.builder()
                                .title("生日")
                                .dataType(Date.class)
                                .extConfigs(createExtConfigs(new DateFormatCfg("yyyy-MM-dd")))
                                .build(),
                        RdColumnDef.builder()
                                .title("年龄")
                                .dataType(Integer.class)
                                .extConfigs(createExtConfigs(new NumFormatCfg("000")))
                                .build(),
                        RdColumnDef.builder()
                                .title("体重")
                                .dataType(Double.class)
                                .build(),
                        RdColumnDef.builder()
                                .title("更新时间")
                                .dataType(Long.class)
                                .extConfigs(createExtConfigs(new DateFormatCfg("yyyy-MM-dd HH:mm:ss"),
                                        new DateZoneIdCfg("UTC-8")))
                                .build(),
                        RdColumnDef.builder()
                                .title("在校生")
                                .dataType(String.class)
                                .build()
                ));

        // ----------写数据方法
        // orgStudentList是保存Student数据的列表

        // 创建一个Excel workbook，并创建一个test-page Sheet页面用于保存数据
        Workbook workbook = ExcelUtil.newXlsxWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");

        // 通过"行数据助理"将数据写入excel page
        RdAssistant.writeObj(page,
                orgStudentList.stream().map(s ->
                        Arrays.asList(s.getName(), s.getBirthday(), s.getAge(), s.getWeight(), s.getUpdateTime(), s.getStudyingText())),
                rdTableDef);

        workbook.write(new FileOutputStream(new File("test.xlsx")));

        // ---------读数据
        // 从文件或者什么流中读入workbook。这里同时支持xlsx或者xls格式
        Workbook readWorkbook = ExcelUtil.openWorkbook(new FileInputStream(new File("test.xlsx")));
        SheetPage readPage = SheetPage.of(readWorkbook, "test-page");

        // 创建用于接收错误信息的recorder
        MessageRecorder messageRecorder = new MessageRecorder();

        // 将page中所有数据读取出来，错误信息记录到recorder中
        List<List<?>> newDataList = RdAssistant.readObjs(readPage, rdTableDef, messageRecorder)
                .collect(Collectors.toList());

        if (messageRecorder.isSuccess()) {
            // 保存数据到数据库
        }
    }
}
