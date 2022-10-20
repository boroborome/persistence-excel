package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.IRdTableDef;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;

@Getter
@AllArgsConstructor
@Builder
public class WriteOptions<D> {
    private String sheetName;
    private IRdTableDef<D, ?> tableDef;
}
