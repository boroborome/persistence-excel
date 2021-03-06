package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.toolkits.manager.ITypeItem;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 配置处理信息。SheetPage用这些信息让各种配置在Excel上生效
 * @param <CT> 对应配置的类型
 */
@Getter
public abstract class RdConfigInfo<CT extends IRdConfig> implements ITypeItem {
    /**
     * 需要处理数据的数据类型
     */
    protected Class<?> type;

    /**
     * 对应配置的类型
     */
    protected Class<CT> configType;

    public RdConfigInfo(Class<CT> configType) {
        this(configType, null);
    }

    public RdConfigInfo(Class<CT> configType, Class<?> type) {
        this.type = type;
        this.configType = configType;
        this.isDataFormat = type != null && type != Void.class && type != Object.class;
    }

    /**
     * 标示当前配置是否属于格式配置<br>
     *     一个Cell上只能有一个格式配置，因此，多个格式配置之间互相冲突。使用的时候只有优先级最高的生效<br>
     *     生效顺序：Cell上配置、Column配置、Sheet上配置。分别对应某次写入是特别指定的配置，字段上的注解，类上的注解与Page上的配置。
     */
    protected boolean isDataFormat;

    /**
     * 将这个配置应用到CellStyle上
     * @param cellStyle 等待配置的cellStyle
     * @param rdConfig 需要配置到cellStyle上的配置信息
     * @param bsc 包含当前单元格信息的一些上下文
     */
    public abstract void buildStyle(CellStyle cellStyle, CT rdConfig, BuildStyleContext bsc);
}
