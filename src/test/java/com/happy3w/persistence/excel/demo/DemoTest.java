package com.happy3w.persistence.excel.demo;

import com.alibaba.fastjson.JSON;
import com.happy3w.persistence.core.rowdata.RdAssistant;
import com.happy3w.persistence.core.rowdata.obj.ObjRdTableDef;
import com.happy3w.persistence.excel.ExcelUtil;
import com.happy3w.persistence.excel.SheetPage;
import com.happy3w.toolkits.message.MessageRecorder;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class DemoTest {
    @Test
    public void should_read_write_by_ann_success() throws IOException {
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

        // 写数据方法
        // orgStudentList是保存Student数据的列表

        // 创建一个Excel workbook，并创建一个test-page Sheet页面用于保存数据
        Workbook workbook = ExcelUtil.newXlsxWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");

        // 重点：通过Student创建一个"行数据表定义"
        ObjRdTableDef<Student> objRdTableDef = ObjRdTableDef.from(Student.class);

        // 通过"行数据助理"将数据写入excel page
        RdAssistant.writeObj(orgStudentList.stream(), page, objRdTableDef);

        workbook.write(new FileOutputStream(new File("test.xlsx")));

        page.locate(0, 0);
        MessageRecorder messageRecorder = new MessageRecorder();
        List<Student> newDataList = RdAssistant.readObjs(page, objRdTableDef, messageRecorder)
                .collect(Collectors.toList());

        Assert.assertEquals(JSON.toJSONString(orgStudentList),
                JSON.toJSONString(newDataList));
    }
}
