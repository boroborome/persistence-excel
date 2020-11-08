package com.happy3w.persistence.excel.style;

import com.happy3w.persistence.core.rowdata.config.DateFormatImpl;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Function;

public class DateFormatStyleBuilder {
    public static void build(CellStyle cellStyle, DateFormatImpl dateFormat, Function<String, Short> formatGetter) {
        cellStyle.setDataFormat(formatGetter.apply(dateFormat.getFormat()));
    }
}
