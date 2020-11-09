package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.IAnnotationRdConfig;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;


@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class FillForegroundColorImpl implements IAnnotationRdConfig<FillForegroundColor> {
    private short color;
    @Override
    public void initBy(FillForegroundColor annotation) {
        this.color = annotation.value();
    }

    @Override
    public void buildContentKey(StringBuilder builder) {
        builder.append(color);
    }
}
