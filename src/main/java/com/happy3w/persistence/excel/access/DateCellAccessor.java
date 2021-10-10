package com.happy3w.persistence.excel.access;

import com.happy3w.java.ext.StringUtils;
import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.config.DateFormatCfg;
import com.happy3w.persistence.core.rowdata.config.DateZoneIdCfg;
import com.happy3w.toolkits.message.MessageRecorderException;
import com.happy3w.toolkits.utils.ZoneIdCache;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

@Getter
@Setter
public class DateCellAccessor implements ICellAccessor<Date> {
    private String nullText;

    @Override
    public void write(Cell cell, Date value, ExtConfigs extConfigs) {
        ZoneId zoneId = getZoneId(extConfigs);

        Calendar calendar = Calendar.getInstance();
        calendar.setTime(value);
        calendar.setTimeZone(TimeZone.getTimeZone(zoneId));
        cell.setCellValue(calendar);
    }

    private ZoneId getZoneId(ExtConfigs extConfigs) {
        DateZoneIdCfg zoneIdConfig = extConfigs.getConfig(DateZoneIdCfg.class);
        String zoneIdStr = zoneIdConfig == null ? ZoneId.systemDefault().getId() :  zoneIdConfig.getZoneId();
        ZoneId zoneId = ZoneIdCache.getZoneId(zoneIdStr);
        return zoneId;
    }

    @Override
    public Date read(Cell cell, Class<?> valueType, ExtConfigs extConfigs, ICellAccessContext context) {
        CellValue cv = context.readCellValue(cell);
        CellType cellType = cv.getCellTypeEnum();
        if (CellType.BLANK.equals(cellType)) {
            return null;
        } else if (cellType == CellType.NUMERIC) {
            ZoneId zoneId = getZoneId(extConfigs);
//            Date cellDate = cell.getDateCellValue();
            Date cellDate = DateUtil.getJavaDate(cv.getNumberValue(), false);
            Instant instant = LocalDateTime.ofInstant(cellDate.toInstant(), ZoneId.systemDefault())
                    .atZone(zoneId)
                    .toInstant();
            return Date.from(instant);
        } else if (!CellType.STRING.equals(cellType)) {
            throw new MessageRecorderException("Can't read date from cell type:" + cellType);
        }

        String strDate = cv.getStringValue();
        if (isNull(strDate)) {
            return null;
        }

        String dateFormat = findDateFormatWithDefault(extConfigs, "yyyy-MM-dd HH:mm:ss");
        try {
            return StringUtils.isEmpty(strDate) ? null : new SimpleDateFormat(dateFormat).parse(strDate);
        } catch (ParseException e) {
            throw new IllegalArgumentException("Failed to parse date:" + strDate, e);
        }
    }

    private boolean isNull(String strDate) {
        return StringUtils.isEmpty(strDate) || (nullText != null && nullText.equals(strDate));
    }

    private String findDateFormatWithDefault(ExtConfigs extConfigs, String defaultFormat) {
        if (extConfigs == null) {
            return defaultFormat;
        }
        DateFormatCfg cfg = extConfigs.getConfig(DateFormatCfg.class);
        if (cfg == null) {
            return defaultFormat;
        }
        return cfg.getFormat();
    }

    @Override
    public Class<Date> getType() {
        return Date.class;
    }
}
