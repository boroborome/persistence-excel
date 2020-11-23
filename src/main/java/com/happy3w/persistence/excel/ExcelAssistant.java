package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.IRdTableDef;
import com.happy3w.persistence.core.rowdata.RdRowIterator;
import com.happy3w.persistence.core.rowdata.RdRowWrapper;
import com.happy3w.toolkits.iterator.EasyIterator;
import com.happy3w.toolkits.iterator.IEasyIterator;
import com.happy3w.toolkits.message.MessageRecorder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.stream.Stream;

@Slf4j
public class ExcelAssistant {

    /**
     * 从当前位置读取数据。当前位置包含Title
     * @param tableDef 行数据定义
     * @param workbook excel的workbook
     * @param messageRecorder 消息记录器
     * @param <D> 行数据的类型
     * @return 以流的形式返回所有行数据
     */
    public static <D> IEasyIterator<RdRowWrapper<D>> readRowsIt(
            IRdTableDef<D, ?> tableDef,
            Workbook workbook,
            MessageRecorder messageRecorder) {
        return EasyIterator.range(0, workbook.getNumberOfSheets())
                .map(index -> SheetPage.of(workbook.getSheetAt(index)))
                .flatMap(page -> RdRowIterator.from(page, tableDef, messageRecorder));
    }

    public static <D> Stream<RdRowWrapper<D>> readRows(
            IRdTableDef<D, ?> tableDef, Workbook workbook, MessageRecorder messageRecorder) {
        return readRowsIt(tableDef, workbook, messageRecorder).stream();
    }
}
