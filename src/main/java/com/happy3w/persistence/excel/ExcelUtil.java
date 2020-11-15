package com.happy3w.persistence.excel;

import com.happy3w.toolkits.message.MessageRecorderException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

public class ExcelUtil {
    public static Workbook openWorkbook(InputStream inputStream) {
        try {
            return WorkbookFactory.create(inputStream);
        } catch (IOException | InvalidFormatException e) {
            throw new MessageRecorderException("Failed to open excel file.", e);
        }
    }

    public static Workbook newXlsWorkbook() {
        return new HSSFWorkbook();
    }

    public static Workbook newXlsxWorkbook() {
        return new XSSFWorkbook();
    }

    public static CellType getCellType(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType type = cell.getCellTypeEnum();
        if (type == CellType.FORMULA) {
            type = cell.getCachedFormulaResultTypeEnum();
        }
        return type;
    }

    // TODO: Date can't be read out
    public static Object readCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType type = getCellType(cell);

        if (type == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (type == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (type == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        } else {
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        }
    }
}
