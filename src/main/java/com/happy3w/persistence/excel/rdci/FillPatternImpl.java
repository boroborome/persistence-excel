package com.happy3w.persistence.excel.rdci;

import com.happy3w.persistence.core.rowdata.IAnnotationRdConfig;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.FillPatternType;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class FillPatternImpl implements IAnnotationRdConfig<FillPattern> {
    private FillPatternType fillPattern;
    @Override
    public void initBy(FillPattern annotation) {
        this.fillPattern = annotation.value();
    }

    @Override
    public void buildContentKey(StringBuilder builder) {
        builder.append(fillPattern);
    }
}
