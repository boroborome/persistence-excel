package com.happy3w.persistence.excel.access;

import com.happy3w.toolkits.convert.TypeConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public interface ICellAccessContext {
    TypeConverter getTypeConverter();
    FormulaEvaluator getFormulaEvaluator();

    default CellValue readCellValue(Cell cell) {
        return getFormulaEvaluator().evaluate(cell);
    }

    default <T> T convert(Object source, Class<T> targetType) {
        return getTypeConverter().convert(source, targetType);
    }
}
