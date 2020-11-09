package com.happy3w.persistence.excel;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@Getter
@Setter
public class CellContext {
    private Object value;
    private Sheet sheet;
    private Workbook workbook;
}
