package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.excel.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;

public class StringCellAccessor implements ICellAccessor<String> {
    @Override
    public void write(Cell cell, String value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public String read(Cell cell, Class<?> valueType, ExtConfigs extConfigs, ICellAccessContext context) {
        CellValue cv = context.readCellValue(cell);
        Object value = ExcelUtil.readCellValue(cv);
        return context.convert(value, String.class);
    }

    @Override
    public Class<String> getType() {
        return String.class;
    }
}
