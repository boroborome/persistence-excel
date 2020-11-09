package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.obj.ObjRdConfigMap;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
@ObjRdConfigMap(FillForegroundColorImpl.class)
public @interface FillForegroundColor {

    /**
     * 调色板中的索引<br>
     *     参考HSSFColor.DARK_RED.index
     * @return
     */
    short value();
}
