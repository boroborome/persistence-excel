package com.happy3w.persistence.excel.writer;

import com.happy3w.persistence.core.rowdata.ExtConfigs;
import com.happy3w.persistence.core.rowdata.config.DateZoneIdImpl;
import com.happy3w.toolkits.utils.ZoneIdCache;
import org.apache.poi.ss.usermodel.Cell;

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
    public Class<Date> getType() {
        return Date.class;
    }
}
