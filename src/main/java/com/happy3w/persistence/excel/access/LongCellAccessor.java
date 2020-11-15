package com.happy3w.persistence.excel.access;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.excel.ExcelUtil;
import com.happy3w.toolkits.convert.TypeConverter;
import org.apache.poi.ss.usermodel.Cell;

public class LongCellAccessor implements ICellAccessor<Long> {
    @Override
    public void write(Cell cell, Long value, ExtConfigs extConfigs) {
        cell.setCellValue(value);
    }

    @Override
    public Long read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        Object value = ExcelUtil.readCellValue(cell);
        if (value instanceof Number) {
            return ((Number) value).longValue();
        } else {
            return TypeConverter.INSTANCE.convert(value, Long.class);
        }
    }

    @Override
    public Class<Long> getType() {
        return Long.class;
    }
}
