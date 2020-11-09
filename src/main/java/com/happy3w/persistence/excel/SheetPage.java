package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.persistence.core.rowdata.page.AbstractWriteDataPage;
import com.happy3w.persistence.core.rowdata.page.IReadDataPage;
import com.happy3w.persistence.excel.access.CellAccessManager;
import com.happy3w.persistence.excel.access.ICellAccessor;
import com.happy3w.persistence.excel.rdci.RdciHolder;
import com.happy3w.toolkits.convert.TypeConverter;
import com.happy3w.toolkits.manager.TypeItemManager;
import com.happy3w.toolkits.utils.Pair;
import com.happy3w.toolkits.utils.PrimitiveTypeUtil;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.MessageFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class SheetPage extends AbstractWriteDataPage<SheetPage> implements IReadDataPage<SheetPage> {
    @Getter
    private final Sheet sheet;

    @Getter
    @Setter
    private TypeConverter valueConverter;

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

    private BuildStyleContext buildStyleContext = new BuildStyleContext();

    public SheetPage(Sheet sheet) {
        this.sheet = sheet;
        buildStyleContext.setSheet(sheet);
        buildStyleContext.setWorkbook(sheet.getWorkbook());
        valueConverter = TypeConverter.INSTANCE.newCopy();
        cellAccessManager = CellAccessManager.INSTANCE.newCopy();
        regRdConfigInfos(RdciHolder.ALL_CONFIG_INFOS);
    }

    public void regRdConfigInfos(List<RdConfigInfo<?, ? extends IRdConfig>> rdConfigInfos) {
        for (RdConfigInfo<?, ? extends IRdConfig> rdConfigInfo : rdConfigInfos) {
            regRdConfigInfo(rdConfigInfo);
        }
    }

    public void regRdConfigInfo(RdConfigInfo<?, ? extends IRdConfig> rdConfigInfo) {
        configTypeToInfo.put(rdConfigInfo.getConfigType(), rdConfigInfo);
        if (rdConfigInfo.isDataFormat()) {
            dataTypeToInfo.registItem(rdConfigInfo);
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

        dataType = PrimitiveTypeUtil.toObjType(dataType);
        ExtConfigs columnConfig = columnConfigs.get(column);
        Pair<Class<?>, Class<? extends IRdConfig>> typeInfos = adjustValueAndFormatType(dataType, extConfigs, columnConfig);
        Class<?> expectValueType = typeInfos.getKey();
        Class<? extends IRdConfig> formatConfigType = typeInfos.getValue();

        ExtConfigs mergedConfig = mergeConfigs(formatConfigType, extConfigs, columnConfig, this.extConfigs);
        ICellAccessor accessor = chooseAccessor(expectValueType);
        Object cellValue = accessor.read(cell, expectValueType, mergedConfig);

        return (D) convertValueToExpectType(cellValue, dataType);
    }

    @Override
    public SheetPage writeValueCfg(Object value, ExtConfigs extConfigs) {
        buildStyleContext.setValue(value);
        Cell cell = ensureCell(row, column);

        ExtConfigs columnConfig = columnConfigs.get(column);
        Pair<Class<?>, Class<? extends IRdConfig>> typeInfos =
                adjustValueAndFormatType(value == null ? null : value.getClass(), extConfigs, columnConfig);
        Class<?> expectValueType = typeInfos.getKey();
        Class<? extends IRdConfig> formatConfigType = typeInfos.getValue();

        ExtConfigs mergedConfig = mergeConfigs(formatConfigType, extConfigs, columnConfig, this.extConfigs);
        buildStyleContext.setExtConfigs(mergedConfig);

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
            configInfo.buildStyle(cellStyle, config, buildStyleContext);
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
}
