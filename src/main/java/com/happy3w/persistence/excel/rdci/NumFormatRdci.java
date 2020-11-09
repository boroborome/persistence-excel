package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.config.NumFormatImpl;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Function;

public class NumFormatRdci extends RdConfigInfo<Number, NumFormatImpl> {
    public NumFormatRdci() {
        super(Number.class, NumFormatImpl.class, true);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, NumFormatImpl rdConfig, Function<String, Short> formatGetter) {
        cellStyle.setDataFormat(formatGetter.apply(rdConfig.getFormat()));
    }
}
