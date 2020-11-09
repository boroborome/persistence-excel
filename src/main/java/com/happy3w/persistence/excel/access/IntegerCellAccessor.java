package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class IntegerCellAccessor implements ICellAccessor<Integer> {
    @Override
    public void write(Cell cell, Integer value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Integer read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
            return null;
        }
        return (int) cell.getNumericCellValue();
    }

    @Override
    public Class<Integer> getType() {
        return Integer.class;
    }
}
