package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.toolkits.manager.ITypeItem;
import com.happy3w.toolkits.utils.TernaryConsumer;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.BiConsumer;
import java.util.function.Function;

@Getter
@AllArgsConstructor
class RdConfigInfo<VT, CT extends IRdConfig> implements ITypeItem<VT> {
    private Class<VT> type;
    private Class<CT> configType;
    private TernaryConsumer<CellStyle, CT, Function<String, Short>> styleBuilder;
    private boolean isDataFormat;
}
