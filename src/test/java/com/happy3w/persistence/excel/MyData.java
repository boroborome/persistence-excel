package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.RdRowWrapper;
import com.happy3w.persistence.core.rowdata.config.DateFormat;
import com.happy3w.persistence.core.rowdata.config.DateZoneId;
import com.happy3w.persistence.core.rowdata.config.NumFormat;
import com.happy3w.persistence.core.rowdata.obj.ObjRdColumn;
import com.happy3w.persistence.core.rowdata.obj.ObjRdPostAction;
import com.happy3w.persistence.excel.rdci.FillForegroundColor;
import com.happy3w.persistence.excel.util.HssfColor;
import com.happy3w.toolkits.message.MessageRecorder;
import lombok.*;

import java.util.Date;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@EqualsAndHashCode
public class MyData {
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
