package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.excel.ExcelUtil;
import com.happy3w.toolkits.convert.TypeConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;

public class DoubleCellAccessor implements ICellAccessor<Double> {
    @Override
    public void write(Cell cell, Double value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Double read(Cell cell, Class<?> valueType, ExtConfigs extConfigs, ICellAccessContext context) {
        CellValue cv = context.readCellValue(cell);
        Object value = ExcelUtil.readCellValue(cv);
        return context.convert(value, Double.class);
    }

    @Override
    public Class<Double> getType() {
        return Double.class;
    }
}
