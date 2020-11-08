package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.persistence.core.rowdata.config.DateFormatImpl;
import com.happy3w.persistence.core.rowdata.config.NumFormatImpl;
import com.happy3w.persistence.core.rowdata.page.AbstractWriteDataPage;
import com.happy3w.persistence.core.rowdata.page.IReadDataPage;
import com.happy3w.persistence.excel.config.DateFormatStyleBuilder;
import com.happy3w.persistence.excel.config.NumFormatStyleBuilder;
import com.happy3w.persistence.excel.writer.ICellAccessor;
import com.happy3w.persistence.excel.writer.CellAccessManager;
import com.happy3w.toolkits.convert.SimpleConverter;
import com.happy3w.toolkits.manager.TypeItemManager;
import com.happy3w.toolkits.message.MessageRecorderException;
import com.happy3w.toolkits.utils.MapBuilder;
import com.happy3w.toolkits.utils.Pair;
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
import java.text.MessageFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

@Slf4j
public class SheetPage extends AbstractWriteDataPage<SheetPage> implements IReadDataPage<SheetPage> {

    /**
     * 默认的格式配置<br>
     *     在一个单元格上，只有一个格式配置会生效，各个格式配置是冲突的
     */
    public static final List<RdConfigInfo<?, ? extends IRdConfig>> DEFAULT_FORMAT_CONFIGS = new ArrayList<>();

    static {
        DEFAULT_FORMAT_CONFIGS.add(new RdConfigInfo<>(Number.class, NumFormatImpl.class, NumFormatStyleBuilder::build, true));
        DEFAULT_FORMAT_CONFIGS.add(new RdConfigInfo<>(Date.class, DateFormatImpl.class, DateFormatStyleBuilder::build, true));
    }

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
    private SimpleConverter valueConverter;

    @Getter
    @Setter
    private TypeItemManager<ICellAccessor> cellAccessManager;

    private Map<Class<? extends IRdConfig>, RdConfigInfo> configTypeToInfo = new HashMap<>();
    private Map<Class<? extends IRdConfig>, RdConfigInfo> dataFormatConfigs = new HashMap<>();
    private TypeItemManager<RdConfigInfo> dataTypeToInfo = new TypeItemManager<>();

    /**
     * 在某一列上的配置
     */
    private Map<Integer, ExtConfigs> columnConfigs = new HashMap<>();

    private Map<String, CellStyle> cellStyleMap = new HashMap<>();
    private Map<String, Short> dataFormatMap = new HashMap<>();

    public SheetPage(Sheet sheet) {
        this.sheet = sheet;
        valueConverter = SimpleConverter.getInstance();
        cellAccessManager = CellAccessManager.INSTANCE.newCopy();
        regRdConfigInfos(DEFAULT_FORMAT_CONFIGS);
    }

    public void regRdConfigInfos(List<RdConfigInfo<?, ? extends IRdConfig>> rdConfigInfos) {
        for (RdConfigInfo<?, ? extends IRdConfig> rdConfigInfo : rdConfigInfos) {
            regRdConfigInfo(rdConfigInfo);
        }
    }

    public void regRdConfigInfo(RdConfigInfo<?, ? extends IRdConfig> rdConfigInfo) {
        configTypeToInfo.put(rdConfigInfo.getConfigType(), rdConfigInfo);
        dataTypeToInfo.registItem(rdConfigInfo);
        if (rdConfigInfo.isDataFormat()) {
            dataFormatConfigs.put(rdConfigInfo.getConfigType(), rdConfigInfo);
        }
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

        ExtConfigs columnConfig = columnConfigs.get(column);
        Pair<Class<?>, Class<? extends IRdConfig>> typeInfos =
                adjustValueAndFormatType(value == null ? null : value.getClass(), extConfigs, columnConfig);
        Class<?> expectValueType = typeInfos.getKey();
        Class<? extends IRdConfig> formatConfigType = typeInfos.getValue();

        ExtConfigs mergedConfig = mergeConfigs(formatConfigType, extConfigs, columnConfig, this.extConfigs);

        Object finalValue = convertValueToExpectType(value, expectValueType);
        if (finalValue != null) {
            ICellAccessor accessor = chooseAccessor(finalValue.getClass());
            accessor.write(cell, finalValue, mergedConfig);
        }

        configCellStyle(cell, mergedConfig);
        column += getColumnSize(cell);

        return this;
    }

    private void configCellStyle(Cell cell, ExtConfigs extConfigs) {
        if (extConfigs.getConfigs().isEmpty()) {
            return;
        }
        String contentKey = extConfigs.createContentKey();
        CellStyle cellStyle = cellStyleMap.computeIfAbsent(contentKey, key -> createCellStyle(extConfigs));
        cell.setCellStyle(cellStyle);
    }

    private CellStyle createCellStyle(ExtConfigs extConfigs) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        for (IRdConfig config : extConfigs.getConfigs().values()) {
            RdConfigInfo<?, IRdConfig> configInfo = configTypeToInfo.get(config.getClass());
            if (configInfo == null) {
                continue;
            }
            configInfo.getStyleBuilder().accept(cellStyle, config, format -> getDataFormat(format));
        }
        return cellStyle;
    }

    private Object convertValueToExpectType(Object value, Class<?> expectValueType) {
        if (expectValueType == null
                || value == null
                || expectValueType.isAssignableFrom(value.getClass())) {
            return value;
        }
        return valueConverter.convert(value, expectValueType);
    }

    private Class<? extends IRdConfig> findFirstDataFormatConfig(ExtConfigs[] candidateConfigs) {
        for (ExtConfigs config : candidateConfigs) {
            if (config == null || config.getConfigs().isEmpty()) {
                continue;
            }

            for (Class<? extends IRdConfig> configType : dataFormatConfigs.keySet()) {
                if (config.getConfig(configType) != null) {
                    return configType;
                }
            }
        }
        return null;
    }

    private Pair<Class<?>, Class<? extends IRdConfig>> adjustValueAndFormatType(Class<?> valueType, ExtConfigs... extConfigs) {
        Class<?> expectValueType = valueType;
        Class<? extends IRdConfig> formatConfigType = findFirstDataFormatConfig(extConfigs);

        if (formatConfigType == null) {
            if (expectValueType != null) {
                RdConfigInfo configInfo = dataTypeToInfo.findItemInheritType(expectValueType);
                if (configInfo != null) {
                    formatConfigType = configInfo.getConfigType();
                }
            }
        } else {
            Class<?> suggestType = configTypeToInfo.get(formatConfigType).getType();
            if (expectValueType == null || !suggestType.isAssignableFrom(expectValueType)) {
                expectValueType = suggestType;
            }
        }
        return new Pair<>(expectValueType, formatConfigType);
    }

    private ExtConfigs mergeConfigs(Class<? extends IRdConfig> formatConfigType, ExtConfigs... extConfigs) {
        ExtConfigs mergedConfig = new ExtConfigs();
        for (int configIndex = extConfigs.length - 1; configIndex >= 0; --configIndex) {
            ExtConfigs extConfig = extConfigs[configIndex];
            if (extConfig == null || extConfig.getConfigs().isEmpty()) {
                continue;
            }
            for (IRdConfig config : extConfig.getConfigs().values()) {
                if (dataFormatConfigs.containsKey(config.getClass())) {
                    if (config.getClass() == formatConfigType) {
                        mergedConfig.regist(config);
                    } else {
                        continue;
                    }
                }
                mergedConfig.regist(config);
            }
        }
        return mergedConfig;
    }

    private ICellAccessor chooseAccessor(Class valueType) {
        ICellAccessor writer = cellAccessManager.findItemInheritType(valueType);
        if (writer == null) {
            throw new UnsupportedOperationException(
                    MessageFormat.format("Unsupported type {0}, no write for it.", valueType));
        }
        return writer;
    }

    public int getColumnSize(Cell cell) {
        for (CellRangeAddress range : cell.getSheet().getMergedRegions()) {
            if (range.containsColumn(cell.getColumnIndex()) && range.containsRow(cell.getRowIndex())) {
                return range.getLastColumn() - range.getFirstColumn() + 1;
            }
        }
        return 1;
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
}
