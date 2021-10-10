package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.excel.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;

public class LongCellAccessor implements ICellAccessor<Long> {
    @Override
    public void write(Cell cell, Long value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Long read(Cell cell, Class<?> valueType, ExtConfigs extConfigs, ICellAccessContext context) {
        CellValue cv = context.readCellValue(cell);
        Object value = ExcelUtil.readCellValue(cv);
        return context.convert(value, Long.class);
    }

    @Override
    public Class<Long> getType() {
        return Long.class;
    }
}
