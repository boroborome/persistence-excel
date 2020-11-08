package com.happy3w.persistence.excel.writer;

import com.happy3w.toolkits.manager.TypeItemManager;

public class CellAccessManager {
    public static final TypeItemManager<ICellAccessor> INSTANCE = new TypeItemManager<ICellAccessor>();

    static {
        INSTANCE.registItem(new StringCellAccessor());
        INSTANCE.registItem(new NumCellAccessor());
        INSTANCE.registItem(new DateCellAccessor());
    }
}
