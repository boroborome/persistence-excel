package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.excel.CellContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

public class FillForegroundColorRdci extends RdConfigInfo<Void, FillForegroundColorImpl> {
    public FillForegroundColorRdci() {
        super(FillForegroundColorImpl.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, FillForegroundColorImpl rdConfig, CellContext cellContext) {
        cellStyle.setFillForegroundColor(rdConfig.getColor());
    }
}
