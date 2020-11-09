package com.happy3w.persistence.excel.access;

import com.happy3w.toolkits.manager.TypeItemManager;

public class CellAccessManager {
    public static final TypeItemManager<ICellAccessor> INSTANCE = new TypeItemManager<ICellAccessor>();

    static {
        INSTANCE.registItem(new IntegerCellAccessor());
        INSTANCE.registItem(new LongCellAccessor());
        INSTANCE.registItem(new StringCellAccessor());
        INSTANCE.registItem(new DoubleCellAccessor());
        INSTANCE.registItem(new DateCellAccessor());
    }
}
