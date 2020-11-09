package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.config.NumFormatImpl;
import com.happy3w.persistence.excel.BuildStyleContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

public class NumFormatRdci extends RdConfigInfo<Number, NumFormatImpl> {
    public NumFormatRdci() {
        super(NumFormatImpl.class, Number.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, NumFormatImpl rdConfig, BuildStyleContext buildStyleContext) {
        short formatId = buildStyleContext.getWorkbook()
                .createDataFormat()
                .getFormat(rdConfig.getFormat());
        cellStyle.setDataFormat(formatId);
    }
}
