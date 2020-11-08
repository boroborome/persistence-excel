package com.happy3w.persistence.excel.style;

import com.happy3w.persistence.core.rowdata.config.NumFormatImpl;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Function;

public class NumFormatStyleBuilder {
    public static void build(CellStyle cellStyle, NumFormatImpl numFormat, Function<String, Short> formatGetter) {
        cellStyle.setDataFormat(formatGetter.apply(numFormat.getFormat()));
    }
}
