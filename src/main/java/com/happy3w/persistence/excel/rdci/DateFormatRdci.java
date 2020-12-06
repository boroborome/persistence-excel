package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.config.DateFormatCfg;
import com.happy3w.persistence.excel.BuildStyleContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Date;

public class DateFormatRdci extends RdConfigInfo<DateFormatCfg> {
    public DateFormatRdci() {
        super(DateFormatCfg.class, Date.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, DateFormatCfg rdConfig, BuildStyleContext bsc) {
        short formatId = bsc.getWorkbook()
                .createDataFormat()
                .getFormat(rdConfig.getFormat());
        cellStyle.setDataFormat(formatId);
    }

}
