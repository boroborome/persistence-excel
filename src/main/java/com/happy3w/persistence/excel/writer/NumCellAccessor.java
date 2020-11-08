package com.happy3w.persistence.excel.writer;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;

public class NumCellAccessor implements ICellAccessor<Number> {
    @Override
    public void write(Cell cell, Number value, ExtConfigs extConfigs) {
        cell.setCellValue(((Number) value).doubleValue());
    }

    @Override
    public Class<Number> getType() {
        return Number.class;
    }
}
