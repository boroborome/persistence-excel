package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.excel.BuildStyleContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

public class FillForegroundColorRdci extends RdConfigInfo<Void, FillForegroundColorImpl> {
    public FillForegroundColorRdci() {
        super(FillForegroundColorImpl.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, FillForegroundColorImpl rdConfig, BuildStyleContext buildStyleContext) {
        cellStyle.setFillForegroundColor(rdConfig.getColor());
    }
}
