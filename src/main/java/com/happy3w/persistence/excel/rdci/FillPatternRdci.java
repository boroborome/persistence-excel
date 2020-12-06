package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.excel.BuildStyleContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

public class FillPatternRdci extends RdConfigInfo<FillPatternCfg> {
    public FillPatternRdci() {
        super(FillPatternCfg.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, FillPatternCfg rdConfig, BuildStyleContext bsc) {
        cellStyle.setFillPattern(rdConfig.getFillPattern());
    }
}
