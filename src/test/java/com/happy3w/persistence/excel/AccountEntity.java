package com.happy3w.persistence.excel;

import com.happy3w.persistence.core.rowdata.obj.ObjRdColumn;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Date;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class AccountEntity {
    @ObjRdColumn("ID")
    private String id;

    @ObjRdColumn("IDCARD")
    private String IdCard;

    @ObjRdColumn("PHONENO")
    private String PhoneNo;

    @ObjRdColumn("CREATEDATE")
    private Date createDate;

    @ObjRdColumn("COMPANY")
    private String company;

    @ObjRdColumn("BANKCARD")
    private String bankCard;

    @ObjRdColumn("BUSIACCOUNT")
    private String busiAccount;

    @ObjRdColumn("LOGINUSER")
    private String loginUser;

    @ObjRdColumn("LOGINPSW")
    private String LoginPsw;

    @ObjRdColumn("TRANPSW")
    private String tranPsw;

    @ObjRdColumn("QUERYPSW")
    private String queryPsw;

    @ObjRdColumn("REMARK")
    private String remark;
}
