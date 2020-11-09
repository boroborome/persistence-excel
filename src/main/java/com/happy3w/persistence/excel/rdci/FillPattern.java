package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.obj.ObjRdConfigMap;
import org.apache.poi.ss.usermodel.FillPatternType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
@ObjRdConfigMap(FillPatternImpl.class)
public @interface FillPattern {
    FillPatternType value();
}
