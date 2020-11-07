package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.page.AbstractWriteDataPage;
import com.happy3w.persistence.core.rowdata.page.IReadDataPage;
import com.happy3w.toolkits.convert.SimpleConverter;
import com.happy3w.toolkits.message.MessageRecorderException;
import com.happy3w.toolkits.utils.MapBuilder;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.MessageFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@Slf4j
public class SheetPage extends AbstractWriteDataPage<SheetPage> implements IReadDataPage<SheetPage> {
    private static final Function<Cell, Object> NULLABLE_NUMBER_READER_FUNCTION = c -> {
        if (CellType.BLANK.equals(c.getCellTypeEnum())) {
            return null;
        }
        return c.getNumericCellValue();
    };

    private static final Function<Cell, Object> NULLABLE_DATE_READER_FUNCTION = c -> {
        if (CellType.BLANK.equals(c.getCellTypeEnum())) {
            return null;
        }
        if (CellType.NUMERIC.equals(c.getCellTypeEnum())) {
            return c.getDateCellValue();
        }
        if (!CellType.STRING.equals(c.getCellTypeEnum())) {
            throw new MessageRecorderException("Can't read date from cell type:" + c.getCellTypeEnum());
        }
        String strDate = c.getStringCellValue().trim();
        Date date = null;
        try {
            date = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(strDate);
        } catch (ParseException e) {
            log.error("Failed to parse date:" + strDate, e);
        }
        return date;
    };

    private static final Map<Class, Function<Cell, Object>> CELL_READER_MAP = MapBuilder
            .<Class, Function<Cell, Object>>of(String.class, (c) -> {
                c.setCellType(CellType.STRING);
                return c.getStringCellValue().trim();
            })
            .and(double.class, c -> c.getNumericCellValue())
            .and(Double.class, NULLABLE_NUMBER_READER_FUNCTION)
            .and(Integer.class, c -> (int) c.getNumericCellValue())
            .and(Long.class, c -> (long) c.getNumericCellValue())
            .and(Date.class, NULLABLE_DATE_READER_FUNCTION)
            .and(BigDecimal.class, c -> new BigDecimal(((XSSFCell) c).getRawValue()))
            .build();

    @Getter
    private final Sheet sheet;

    @Getter
    @Setter
    private SimpleConverter converter;

    private Map<String, CellStyle> cellStyleMap = new HashMap<>();
    private Map<String, Short> dataFormatMap = new HashMap<>();

    public SheetPage(Sheet sheet) {
        this.sheet = sheet;
        converter = SimpleConverter.getInstance();
    }

    @Override
    public String getPageName() {
        return sheet.getSheetName();
    }

    @Override
    public <D> D readValue(int rowIndex, int columnIndex, Class<D> dataType, ExtConfigs extConfigs) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }

        Function<Cell, Object> reader = CELL_READER_MAP.get(dataType);
        if (reader != null) {
            try {
                return (D) reader.apply(cell);
            } catch (Exception e) {
                throw new IllegalArgumentException(
                        MessageFormat.format("Failed to read cell value at row:{0}, column:{1}, dataType:{2}.",
                        rowIndex, columnIndex, dataType), e);
            }
        }
        cell.setCellType(CellType.STRING);
        String strValue = cell.getStringCellValue().trim();
        Object newValue = SimpleConverter.getInstance().convert(strValue, dataType);
        return (D) newValue;
    }

    @Override
    public SheetPage writeValueCfg(Object value, ExtConfigs extConfigs) {
        Cell cell = ensureCell(row, column);
        CellStyle columnStyle = getColumnStyle(column);
        if (value instanceof Date) {
            CellStyle cellStyle = getCurrentCellDateStyle(extConfigs);
            writeCellDate(cell, (Date) value, new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"), cellStyle);
        } else if (value instanceof Number) {
            CellStyle cellStyle = getCurrentCellNumStyle(extConfigs);
            if (cellStyle != null) {
                cell.setCellStyle(cellStyle);
            }
            cell.setCellValue(((Number) value).doubleValue());
        } else {
            CellStyle cellStyle = getColumnStyle(column);
            if (cellStyle != null) {
                cell.setCellStyle(cellStyle);
            }
            if (value instanceof List) {
                List<String> strList = (List<String>) ((List) value).stream()
                        .map(item -> String.valueOf(item))
                        .collect(Collectors.toList());
                cell.setCellValue(String.join(",", strList));
            } else if (value == null) {
                cell.setCellStyle(columnStyle);
            } else {
                cell.setCellStyle(columnStyle);
                cell.setCellValue(String.valueOf(value));
            }
        }

        column += getColumnSize(cell);

        return this;
    }

    public int getColumnSize(Cell cell) {
        for (CellRangeAddress range : cell.getSheet().getMergedRegions()) {
            if (range.containsColumn(cell.getColumnIndex()) && range.containsRow(cell.getRowIndex())) {
                return range.getLastColumn() - range.getFirstColumn() + 1;
            }
        }
        return 1;
    }

    private void writeCellDate(Cell cell, Date date, DateFormat dateFormat, CellStyle dateCellStyle) {
        Calendar calendar = Calendar.getInstance(dateFormat.getTimeZone());
        try {
            Date roundDate = dateFormat.parse(dateFormat.format(date));
            calendar.setTime(roundDate);
            cell.setCellValue(calendar);
            cell.setCellStyle(dateCellStyle);
        } catch (ParseException e) {
            log.error("Failed to parse date", e);
        }
    }

    private Cell ensureCell(int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    public CellStyle getColumnStyle(int column) {
        CellStyle columnStyle = sheet.getColumnStyle(column);
        return (columnStyle == null || columnStyle.getIndex() == 0)
                ? null
                : columnStyle;
    }

    private CellStyle getCurrentCellDateStyle(ExtConfigs newConfigs) {
        if (noDateFormat(newConfigs)) {
            CellStyle columnStyle = getColumnStyle(column);
            if (columnStyle != null) {
                return columnStyle;
            } else if (noDateFormat(this.extConfigs)) {
                return null;
            }
            return getStyleByPattern("yyyy-MM-dd HH:mm:ss");
        }
        return getStyleByPattern("yyyy-MM-dd HH:mm:ss");
    }

    private boolean noDateFormat(ExtConfigs extConfigs) {
        return true;
    }

    private CellStyle getStyleByPattern(String datePattern) {
        CellStyle style = cellStyleMap.get(datePattern);
        if (style == null) {
            cellStyle(datePattern, datePattern);
            style = cellStyleMap.get(datePattern);
        }
        return style;
    }

    private CellStyle getCurrentCellNumStyle(ExtConfigs extConfigs) {
        return null;
    }


    public static SheetPage of(Sheet sheet) {
        return new SheetPage(sheet);
    }

    public static SheetPage of(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        return new SheetPage(sheet);
    }

    public SheetPage mergeCell(int rowSize, int columnSize) {
        if (columnSize > 1 || rowSize > 1) {
            sheet.addMergedRegion(new CellRangeAddress(row, row + rowSize - 1,
                    column, column + columnSize - 1));
        }

        return this;
    }

    public short getDataFormat(String dataFormat) {
        Short formatId = dataFormatMap.get(dataFormat);
        if (formatId == null) {
            short df = sheet.getWorkbook().createDataFormat().getFormat(dataFormat);
            formatId = df;
            dataFormatMap.put(dataFormat, formatId);
        }
        return formatId;
    }

    public SheetPage cellStyle(String styleName, String dateFormat) {
        CellStyle cellStyle = cellStyleMap.get(styleName);
        if (cellStyle == null) {
            cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyleMap.put(styleName, cellStyle);
        }

        cellStyle.setDataFormat(getDataFormat(dateFormat));

        return this;
    }
}
