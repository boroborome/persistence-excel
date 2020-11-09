package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.config.DateFormatImpl;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Date;
import java.util.function.Function;

public class DateFormatRdci extends RdConfigInfo<Date, DateFormatImpl> {
    public DateFormatRdci() {
        super(Date.class, DateFormatImpl.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, DateFormatImpl rdConfig, Function<String, Short> formatGetter) {
        cellStyle.setDataFormat(formatGetter.apply(rdConfig.getFormat()));
    }
}
