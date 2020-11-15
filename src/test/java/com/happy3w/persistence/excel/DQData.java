package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.obj.ObjRdColumn;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

@Getter
@Setter
public class DQData {
    @ObjRdColumn("所属交易卡/账号")
    private String account;
    @ObjRdColumn("预计本息合计")
    private String balance;
    @ObjRdColumn("本金")
    private String total;
    private String netValue = "1.01";
    private String costPrice="1";
    @ObjRdColumn("币种")
    private String currency;
    @ObjRdColumn("钞汇")
    private String cashOrExchange;
    @ObjRdColumn("SysTime")
    private Date sysTime;
    @ObjRdColumn("到期日")
    private Date maturityDate;
    @ObjRdColumn("存期")
    private String period;
}
