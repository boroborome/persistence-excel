package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class DoubleCellAccessor implements ICellAccessor<Double> {
    @Override
    public void write(Cell cell, Double value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Double read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
            return null;
        }
        return cell.getNumericCellValue();
    }

    @Override
    public Class<Double> getType() {
        return Double.class;
    }
}
