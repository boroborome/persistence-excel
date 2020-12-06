package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.config.NumFormatCfg;
import com.happy3w.persistence.excel.BuildStyleContext;
import com.happy3w.persistence.excel.RdConfigInfo;
import org.apache.poi.ss.usermodel.CellStyle;

public class NumFormatRdci extends RdConfigInfo<NumFormatCfg> {
    public NumFormatRdci() {
        super(NumFormatCfg.class, Number.class);
    }

    @Override
    public void buildStyle(CellStyle cellStyle, NumFormatCfg rdConfig, BuildStyleContext bsc) {
        short formatId = bsc.getWorkbook()
                .createDataFormat()
                .getFormat(rdConfig.getFormat());
        cellStyle.setDataFormat(formatId);
    }
}
