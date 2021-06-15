package com.happy3w.persistence.excel;

import com.happy3w.persistence.excel.access.DateCellAccessor;
import com.happy3w.persistence.excel.access.ICellAccessor;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import java.util.Date;

import static org.junit.Assert.*;

public class ExcelAssistantTest {

    @Test
    public void should_config_null_text_success() {
        Workbook workbook = ExcelUtil.newXlsxWorkbook();
        SheetPage page = SheetPage.of(workbook, "test-page");
        DateCellAccessor accessor = (DateCellAccessor) page.getCellAccessManager()
                .findByType(Date.class);

        Assert.assertNotNull(accessor);
    }
}