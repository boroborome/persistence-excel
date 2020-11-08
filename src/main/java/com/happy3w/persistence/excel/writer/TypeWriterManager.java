package com.happy3w.persistence.excel.writer;

import com.happy3w.toolkits.manager.TypeItemManager;

public class TypeWriterManager {
    public static final TypeItemManager<ICellTypeWriter> INSTANCE = new TypeItemManager<ICellTypeWriter>();

    static {
        INSTANCE.registItem(new StringCellTypeWriter());
        INSTANCE.registItem(new NumCellTypeWriter());
        INSTANCE.registItem(new DateCellTypeWriter());
    }
}
