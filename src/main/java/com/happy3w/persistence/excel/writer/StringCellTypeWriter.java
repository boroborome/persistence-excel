package com.happy3w.persistence.excel.writer;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;

public class StringCellTypeWriter implements ICellTypeWriter<String>{
    @Override
    public void write(Cell cell, String value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Class<String> getType() {
        return String.class;
    }
}
