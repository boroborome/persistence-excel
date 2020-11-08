package com.happy3w.persistence.excel.writer;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.config.DateFormatImpl;
import com.happy3w.persistence.core.rowdata.config.DateZoneIdImpl;
import com.happy3w.toolkits.message.MessageRecorderException;
import com.happy3w.toolkits.utils.ZoneIdCache;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

public class DateCellAccessor implements ICellAccessor<Date> {
    @Override
    public void write(Cell cell, Date value, ExtConfigs extConfigs) {
        DateZoneIdImpl zoneIdConfig = extConfigs.getConfig(DateZoneIdImpl.class);
        String zoneIdStr = zoneIdConfig == null ? ZoneId.systemDefault().getId() :  zoneIdConfig.getZoneId();
        ZoneId zoneId = ZoneIdCache.getZoneId(zoneIdStr);

        Calendar calendar = Calendar.getInstance();
        calendar.setTime(value);
        calendar.setTimeZone(TimeZone.getTimeZone(zoneId));
        cell.setCellValue(calendar);
    }

    @Override
    public Date read(Cell cell, Class<?> valueType, ExtConfigs extConfigs) {
        if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
            return null;
        }
        if (CellType.NUMERIC.equals(cell.getCellTypeEnum())) {
            return cell.getDateCellValue();
        }
        if (!CellType.STRING.equals(cell.getCellTypeEnum())) {
            throw new MessageRecorderException("Can't read date from cell type:" + cell.getCellTypeEnum());
        }

        String dateFormat = extConfigs.getConfig(DateFormatImpl.class).getFormat();
        String strDate = cell.getStringCellValue().trim();
        Date date = null;
        try {
            date = new SimpleDateFormat(dateFormat).parse(strDate);
        } catch (ParseException e) {
            throw new IllegalArgumentException("Failed to parse date:" + strDate, e);
        }
        return date;
    }

    @Override
    public Class<Date> getType() {
        return Date.class;
    }
}
