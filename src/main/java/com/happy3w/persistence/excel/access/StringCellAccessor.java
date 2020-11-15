package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class StringCellAccessor implements ICellAccessor<String> {
    @Override
    public void write(Cell cell, String value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public String read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (cell.getCellTypeEnum() != CellType.STRING) {
            cell.setCellType(CellType.STRING);
        }
        return cell.getStringCellValue();
    }

    @Override
    public Class<String> getType() {
        return String.class;
    }
}
