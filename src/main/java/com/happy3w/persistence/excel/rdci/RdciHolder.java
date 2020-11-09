package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.persistence.excel.RdConfigInfo;

import java.util.ArrayList;
import java.util.List;

public class RdciHolder {
    /**
     * 默认的行数据配置信息<br>
     *     这里信息修改后，所有新创建的SheetPage会使用这个信息，之前创建的不会变化
     */
    public static final List<RdConfigInfo<?, ? extends IRdConfig>> ALL_CONFIG_INFOS = new ArrayList<>();

    static {
        ALL_CONFIG_INFOS.add(new NumFormatRdci());
        ALL_CONFIG_INFOS.add(new DateFormatRdci());
        ALL_CONFIG_INFOS.add(new FillForegroundColorRdci());
    }
}
