package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.config.DateFormatImpl;
import com.happy3w.persistence.excel.BuildStyleContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Date;

public class DateFormatRdci extends RdConfigInfo<Date, DateFormatImpl> {
    public DateFormatRdci() {
        super(DateFormatImpl.class, Date.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, DateFormatImpl rdConfig, BuildStyleContext bsc) {
        short formatId = bsc.getWorkbook()
                .createDataFormat()
                .getFormat(rdConfig.getFormat());
        cellStyle.setDataFormat(formatId);
    }

}
