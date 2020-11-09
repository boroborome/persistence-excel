package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class LongCellAccessor implements ICellAccessor<Long> {
    @Override
    public void write(Cell cell, Long value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Long read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
            return null;
        }
        return (long) cell.getNumericCellValue();
    }

    @Override
    public Class<Long> getType() {
        return Long.class;
    }
}
