package com.happy3w.persistence.excel.writer;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.toolkits.manager.ITypeItem;
import org.apache.poi.ss.usermodel.Cell;

public interface ICellAccessor<T> extends ITypeItem<T> {
    void write(Cell cell, T value, ExtConfigs extConfigs);
    T read(Cell cell, Class<?> valueType, ExtConfigs extConfigs);
}
