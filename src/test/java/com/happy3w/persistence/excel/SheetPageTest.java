package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.RdAssistant;
import com.happy3w.persistence.core.rowdata.obj.ObjRdColumn;
import com.happy3w.persistence.core.rowdata.obj.ObjRdTableDef;
import com.happy3w.toolkits.message.MessageRecorder;
import junit.framework.TestCase;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class SheetPageTest extends TestCase {

    public void test_read_write_success_when_normal() throws IOException {
        List<MyData> dataList = Arrays.asList(new MyData("Tom", 12),
                new MyData("张三", 21));

        Workbook workbook = ExcelUtil.newXlsWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");

        ObjRdTableDef<MyData> objRdTableDef = ObjRdTableDef.from(MyData.class);
        RdAssistant.writeObj(dataList.stream(), page, objRdTableDef);

//        File excelFile = File.createTempFile("temp-excel", ".xlsx");
//        workbook.write(new FileOutputStream(excelFile));

        page.locate(0, 0);
        MessageRecorder messageRecorder = new MessageRecorder();
        List<MyData> readDatas = RdAssistant.readObjs(objRdTableDef, page, messageRecorder)
                .collect(Collectors.toList());

        Assert.assertEquals(dataList.size(), readDatas.size());
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    public static class MyData {
        @ObjRdColumn("名字")
        private String name;

        @ObjRdColumn("年龄")
        private int age;
    }
}
