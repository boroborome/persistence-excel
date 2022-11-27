package com.happy3w.persistence.excel;

import com.happy3w.persistence.excel.access.DateCellAccessor;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.util.Date;

public class ExcelAssistantTest {

    @Test
    public void should_config_null_text_success() {
        Workbook workbook = ExcelUtil.newXlsxWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");
        DateCellAccessor accessor = (DateCellAccessor) page.getCellAccessManager()
                .findByType(Date.class);

        Assertions.assertNotNull(accessor);
    }
}