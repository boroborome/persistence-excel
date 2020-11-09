package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@Getter
@Setter
public class BuildStyleContext {
    private Sheet sheet;
    private Workbook workbook;
    private Object value;
    private ExtConfigs extConfigs;
}
