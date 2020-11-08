package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class NumCellAccessor implements ICellAccessor<Number> {
    @Override
    public void write(Cell cell, Number value, ExtConfigs extConfigs) {
        cell.setCellValue(((Number) value).doubleValue());
    }

    @Override
    public Number read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
            return null;
        }
        return cell.getNumericCellValue();
    }

    @Override
    public Class<Number> getType() {
        return Number.class;
    }
}
