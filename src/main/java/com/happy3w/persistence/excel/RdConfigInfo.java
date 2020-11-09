package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.IRdConfig;
import com.happy3w.toolkits.manager.ITypeItem;
import com.happy3w.toolkits.utils.TernaryConsumer;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * 配置处理信息。SheetPage用这些信息让各种配置在Excel上生效
 * @param <VT> 需要处理数据的数据类型
 * @param <CT> 对应配置的类型
 */
@Getter
@AllArgsConstructor
class RdConfigInfo<VT, CT extends IRdConfig> implements ITypeItem<VT> {
    /**
     * 需要处理数据的数据类型
     */
    private Class<VT> type;

    /**
     * 对应配置的类型
     */
    private Class<CT> configType;

    /**
     * 将这个而配置应用到CellStyle上
     */
    private TernaryConsumer<CellStyle, CT, Function<String, Short>> styleBuilder;

    /**
     * 标示当前配置是否属于格式配置<br>
     *     一个Cell上只能有一个格式配置，因此，多个格式配置之间互相冲突。使用的时候只有优先级最高的生效<br>
     *     生效顺序：Cell上配置、Column配置、Sheet上配置。分别对应某次写入是特别指定的配置，字段上的注解，类上的注解与Page上的配置。
     */
    private boolean isDataFormat;
}
