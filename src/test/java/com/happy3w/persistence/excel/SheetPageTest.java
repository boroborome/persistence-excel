package com.happy3w.persistence.excel;

import com.alibaba.fastjson.JSON;
import com.happy3w.persistence.core.rowdata.RdAssistant;
import com.happy3w.persistence.core.rowdata.RdRowWrapper;
import com.happy3w.persistence.core.rowdata.config.DateFormat;
import com.happy3w.persistence.core.rowdata.config.DateZoneId;
import com.happy3w.persistence.core.rowdata.config.NumFormat;
import com.happy3w.persistence.core.rowdata.obj.ObjRdColumn;
import com.happy3w.persistence.core.rowdata.obj.ObjRdPostAction;
import com.happy3w.persistence.core.rowdata.obj.ObjRdTableDef;
import com.happy3w.persistence.excel.rdci.FillForegroundColor;
import com.happy3w.persistence.excel.util.HssfColor;
import com.happy3w.toolkits.convert.SimpleConverter;
import com.happy3w.toolkits.message.MessageRecorder;
import junit.framework.TestCase;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class SheetPageTest extends TestCase {

    public void test_read_write_success_when_normal() throws IOException {
        List<MyData> orgDataList = Arrays.asList(
                MyData.builder().name("Tom")
                        .age(12)
                        .enabled(true)
                        .favoriteDate(SimpleConverter.getInstance().convert("2020-10-10 23:00:00", Date.class).getTime())
                        .build(),
                MyData.builder().name("张三")
                        .age(21)
                        .birthday(SimpleConverter.getInstance().convert("2020-10-10 23:00:00", Date.class))
                        .build());

        Workbook workbook = ExcelUtil.newXlsxWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");

        ObjRdTableDef<MyData> objRdTableDef = ObjRdTableDef.from(MyData.class);
        RdAssistant.writeObj(orgDataList.stream(), page, objRdTableDef);

        page.locate(0, 0);
        MessageRecorder messageRecorder = new MessageRecorder();
        List<MyData> newDataList = RdAssistant.readObjs(page, objRdTableDef, messageRecorder)
                .collect(Collectors.toList());

        Assert.assertEquals(JSON.toJSONString(orgDataList),
                JSON.toJSONString(newDataList));
    }

    public void test_read_formula_success() {
        Workbook workbook = ExcelUtil.openWorkbook(SheetPage.class.getResourceAsStream("/formula-excel.xlsx"));
        SheetPage page = SheetPage.of(workbook, "Sheet1");

        ObjRdTableDef<MyData> objRdTableDef = ObjRdTableDef.from(MyData.class);
        MessageRecorder messageRecorder = new MessageRecorder();
        List<MyData> newDataList = RdAssistant.readObjs(page, objRdTableDef, messageRecorder)
                .collect(Collectors.toList());

        Assert.assertEquals("[{\"age\":15,\"enabled\":true,\"enabledText\":\"true\",\"name\":\"Tom\"}]",
                JSON.toJSONString(newDataList));
    }

    public void test_read_formula_systime_success() {
        Workbook workbook = ExcelUtil.openWorkbook(SheetPage.class.getResourceAsStream("/dq.xlsx"));
        SheetPage page = SheetPage.of(workbook, "定期");

        ObjRdTableDef<DQData> objRdTableDef = ObjRdTableDef.from(DQData.class);
        MessageRecorder messageRecorder = new MessageRecorder();
        List<DQData> newDataList = RdAssistant.readObjs(page, objRdTableDef, messageRecorder)
                .collect(Collectors.toList());

        Assert.assertEquals("[{\"age\":15,\"enabled\":false,\"enabledText\":\"false\",\"name\":\"Tom\"}]",
                JSON.toJSONString(newDataList));
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    @EqualsAndHashCode
    public static class MyData {
        @FillForegroundColor(HssfColor.RED)
        @ObjRdColumn(value = "名字")
        private String name;

        @ObjRdColumn(value = "年龄", required = false)
        @NumFormat("000")
        private int age;

        @ObjRdColumn(value = "在校生", getter = "getEnabledText", setter = "setEnabledText")
        private boolean enabled;

        @ObjRdColumn("生日")
        @DateFormat("yyyy-MM-dd HH:mm:ss")
        private Date birthday;

        @ObjRdColumn("Favorite Date")
        @DateFormat("yyyy-MM-dd HH:mm:ss")
        @DateZoneId("UTC-8")
        private Long favoriteDate;

        @ObjRdPostAction
        public void postInit(RdRowWrapper<MyData> data, MessageRecorder recorder) {

        }

        public String getEnabledText() {
            return Boolean.toString(enabled);
        }

        public void setEnabledText(String enabled, RdRowWrapper<MyData> data, MessageRecorder recorder) {
            this.enabled = Boolean.parseBoolean(enabled);
        }
    }
}
