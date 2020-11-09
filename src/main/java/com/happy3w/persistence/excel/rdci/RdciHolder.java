package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.persistence.excel.RdConfigInfo;

import java.util.ArrayList;
import java.util.List;

public class RdciHolder {
    /**
     * 默认的配置处理信息<br>
     */
    public static final List<RdConfigInfo<?, ? extends IRdConfig>> ALL_CONFIG_INFOS = new ArrayList<>();

    static {
        ALL_CONFIG_INFOS.add(new NumFormatRdci());
        ALL_CONFIG_INFOS.add(new DateFormatRdci());
    }
}
