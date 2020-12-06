package com.happy3w.persistence.excel.demo;

import com.happy3w.persistence.core.rowdata.RdRowWrapper;
import com.happy3w.persistence.core.rowdata.config.DateFormat;
import com.happy3w.persistence.core.rowdata.config.DateZoneId;
import com.happy3w.persistence.core.rowdata.config.NumFormat;
import com.happy3w.persistence.core.rowdata.obj.ObjRdColumn;
import com.happy3w.persistence.core.rowdata.obj.ObjRdPostAction;
import com.happy3w.persistence.excel.rdci.FillForegroundColor;
import com.happy3w.persistence.excel.util.HssfColor;
import com.happy3w.toolkits.message.MessageRecorder;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Date;

@Builder
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@NumFormat("#.00")                          // 配置默认数字格式为固定显示两位小数
public class Student {
    @ObjRdColumn(value = "名字")             // 配置这个字段在文件中的列头名称
    @FillForegroundColor(HssfColor.RED)     // 配置在导出Excel时使用红色背景色（这个在库persistence-excel中）
    private String name;

    @ObjRdColumn("生日")
    @DateFormat("yyyy-MM-dd")               // 配置使用的时间格式
    private Date birthday;

    @ObjRdColumn(value = "年龄", required = false) // 年龄不是必填项
    @NumFormat("000")                       // 年龄显示不需要小数
    private Integer age;

    @ObjRdColumn(value = "体重")    			// 这里没有配置数字格式，使用前面配置的默认格式，两位小数显示
    private double weight;

    @ObjRdColumn("更新时间")
    @DateFormat("yyyy-MM-dd HH:mm:ss")
    @DateZoneId("UTC-8")                    // 配置读写文件时使用的时区
    private long updateTime;

    // 在校生信息需要经过转换才能变成boolean，通过配置的getter和setter转换
    @ObjRdColumn(value = "在校生", getter = "getStudyingText", setter = "setStudyingText")
    private boolean studying;

    /**
     * 配置数据从文件加载后需要额外做的一些操作。比如年龄必须大于0，小于100的检测；名字可能带有不需要的前缀，需要去掉。
     * ObjRdPostAction对被注解的方法名称、参数个数、参数顺序都没有要求，但一个对象只能有一个postAction。工具根据需要自动注入
     * @param data 刚刚解析数据使用的行信息，包括page name，行数等信息
     * @param recorder 如果有需要返回给用户的消息，通过这个recorder记录下来
     */
    @ObjRdPostAction
    public void postInit(RdRowWrapper<Student> data, MessageRecorder recorder) {
        if (age != null && (age < 0 || age > 100)) {
            recorder.appendError("Wrong age:{0}", age);
        }
        if (name.startsWith("Name:")) {
            name = name.substring(5);
        }
    }

    public String getStudyingText() {
        return studying ? "在校" : "毕业";
    }

    // 列头注册的setter方法可以带有两个额外的参数，属性值必须在第一位，其他参数数量和顺序没有要求，工具自动注入
    public void setStudyingText(String studyingText, RdRowWrapper<Student> data, MessageRecorder recorder) {
        this.studying = "在校".equals(studyingText);
    }
}
