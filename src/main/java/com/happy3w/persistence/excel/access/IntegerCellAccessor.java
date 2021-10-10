package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.excel.ExcelUtil;
import com.happy3w.toolkits.convert.TypeConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;

public class IntegerCellAccessor implements ICellAccessor<Integer> {
    @Override
    public void write(Cell cell, Integer value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Integer read(Cell cell, Class<?> valueType, ExtConfigs extConfigs, ICellAccessContext context) {
        CellValue cv = context.readCellValue(cell);
        Object value = ExcelUtil.readCellValue(cv);
        return context.convert(value, Integer.class);
    }

    @Override
    public Class<Integer> getType() {
        return Integer.class;
    }
}
