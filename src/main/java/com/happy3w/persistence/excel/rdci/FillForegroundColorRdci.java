package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Function;

public class FillForegroundColorRdci extends RdConfigInfo<Void, FillForegroundColorImpl> {
    public FillForegroundColorRdci() {
        super(Void.class, FillForegroundColorImpl.class, false);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, FillForegroundColorImpl rdConfig, Function<String, Short> formatGetter) {
        cellStyle.setFillForegroundColor(rdConfig.getColor());
    }
}
