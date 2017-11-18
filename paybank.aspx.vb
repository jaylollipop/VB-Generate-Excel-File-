Imports System.Data

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports OfficeOpenXml

Imports System.IO
Partial Class Default2
    Inherits System.Web.UI.Page

    Dim dbARY As New ClassDBManager("AREEYA")
    Dim dbERP As New ClassDBManager("ERP")
    Dim strSQL, strBack As New StringBuilder
    Dim Thai_Month_Array() As String = {"", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"}
    Dim mDay, mMonth, mYear, mDays, mMonths, mYears As String
    Dim m_exp_amt_all As Decimal
    Dim m_exp_amt As Decimal = 0
    Dim m_exp_amt2 As Decimal = 0
    Dim m_exp_amt3 As Decimal = 0
    Dim m_exp_amt4 As Decimal = 0
    Dim m_exp_amt5 As Decimal = 0
    Dim m_exp_amt6 As Decimal = 0
    Dim m_exp_amt7 As Decimal = 0
    Dim m_exp_amt8 As Decimal = 0
    Dim paybank As String = ""

    'Dim Thai_Month_Array() As String = {"มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"}
    Dim projID As String = ""
    Dim strHtml As String = ""

    Dim Contract As String = ""
    Dim proj As String = ""
    Dim Now_Dates As String = ""
    Dim PD_Code As String = "AAA"
    Dim BD_CODE As String = ""
    Dim FL_CODE As String = ""
    Dim Projcode As String = ""
    Dim CONTRACTNAME As String = " "
    Dim namepro As String = ""
    Dim landWAH As Decimal = 0
    Dim PROPERTYSALESID As String = ""
    Dim totalE As Decimal = 0
    Dim totalB As Decimal = 0
    Dim m_home_area As Decimal = 0
    Dim m_home_price As Decimal = 0
    Dim m_depre As Decimal = 0
    Dim mTransYear As Decimal = 0
    Dim NOW_DAY As Integer = 0
    Dim NOW_MONTH As Integer = 0
    Dim NOW_YEAR As Integer = 0
    Dim D_TRAN As Integer = 0
    Dim M_TRAN As Integer = 0
    Dim Y_TRAN As Integer = 0
    Dim PDCODE As String = " "
    Dim UNITPRICE As Decimal = 0
    Dim PROMOTIONAMOUNT As Decimal = 0
    Dim BASESALES As Decimal = 0
    Dim CashPromotionAmt As Decimal = 0
    Dim promotionlast As Decimal = 0
    Dim now_date As String
    Dim a As Decimal = 100
    Dim b As Decimal = 1

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Get_PDF()
    End Sub
    Private Sub Get_PDF()

        Dim template As FileInfo = New FileInfo(Server.MapPath("PDF/Paybank/paybank.xlsx"))
        Dim package As ExcelPackage = New ExcelPackage(template, True)



        paybank = Request.QueryString("P_SEALID")
        Dim sql As New StringBuilder()
        Dim STATUS As Integer
        strSQL.Remove(0, strSQL.Length)
        strSQL.Append("Select PROPERTYSALESTATUS ")
        strSQL.Append("from IVZ_PROPERTYSALESTABLE ")
        strSQL.Append("where PROPERTYSALESID = '" & paybank & "' ")
        dbERP.sql = strSQL.ToString
        Dim dtsatatus As DataTable = dbERP.QueryDataTable()
        If dtsatatus.Rows.Count >= 1 Then
            STATUS = dtsatatus.Rows(0).Item("PROPERTYSALESTATUS")
        End If

        If STATUS = 0 Or STATUS = 6 Or STATUS = 7 Then
            Response.Write("ไม่สามาพิมพ์ได้ เนื่องจาก สถานนะเป็น Open Order หรือ Sold Out หรือ Calceled ")
        Else

            'mMonth = ndate.Text.Substring(3, 2)
            'mYear = ndate.Text.Substring(6, 4)
            'fromDate = mYear & "/" & mMonth & "/" & mDay

            'Dim Thai_Month_Array() As String = {"มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"}
            'Dim projID As String = ""
            'Dim strHtml As String = ""

            'Dim Contract As String = ""
            'Dim proj As String = ""
            'Dim Now_Dates As String = ""
            'Dim PD_Code As String = "AAA"
            'Dim BD_CODE As String = ""
            'Dim FL_CODE As String = ""
            'Dim Projcode As String = ""
            'Dim CONTRACTNAME As String = " "
            'Dim namepro As String = ""
            'Dim landWAH As Decimal = 0
            'Dim PROPERTYSALESID As String = ""
            'Dim totalE As Decimal = 0
            'Dim totalB As Decimal = 0
            'Dim m_home_area As Decimal = 0
            'Dim m_home_price As Decimal = 0
            'Dim m_depre As Decimal = 0
            'Dim mTransYear As Decimal = 0
            'Dim NOW_DAY As Integer = 0
            'Dim NOW_MONTH As Integer = 0
            'Dim NOW_YEAR As Integer = 0
            'Dim D_TRAN As Integer = 0
            'Dim M_TRAN As Integer = 0
            'Dim Y_TRAN As Integer = 0
            'Dim PDCODE As String = " "
            'Dim UNITPRICE As Decimal = 0
            'Dim PROMOTIONAMOUNT As Decimal = 0
            'Dim BASESALES As Decimal = 0
            'Dim CashPromotionAmt As Decimal = 0
            'Dim promotionlast As Decimal = 0
            'Dim now_date As String
            'Dim a As Decimal = 100
            'Dim b As Decimal = 1
            paybank = Request.QueryString("P_SEALID")
            'Dim sql As New StringBuilder()

            'Dim STATUS As Integer
            'strSQL.Remove(0, sql.Length)
            'strSQL.Append("Select PROPERTYSALESTATUS ")
            'strSQL.Append("from IVZ_PROPERTYSALESTABLE ")
            'strSQL.Append("where PROPERTYSALESID = '" & paybank & "' ")
            'dbERP.sql = strSQL.ToString
            'Dim dtsatatus As DataTable = dbERP.QueryDataTable()
            'If dbERP.RowsAffected >= 1 Then
            '    STATUS = dtsatatus.Rows(0).Item("PROPERTYSALESTATUS")
            'End If
            'If STATUS = 0 Or STATUS = 6 Or STATUS = 7 Then
            '    Response.Write("ไม่สามาพิมพ์ได้ เนื่องจาก สถานนะเป็น Open Order หรือ Sold Out หรือ Calceled ")
            'Else



            'Dim sql As New StringBuilder()



            If paybank <> "" Then
                strSQL.Remove(0, strSQL.Length)
                strSQL.Append(" SELECT A.SalesResponse, A.PROPERTYSALESID,C.ITEMID,A.PROJID,SUBSTRING(B.SUBPROJID,14,1)AS BD_CODE,SUBSTRING(B.SUBPROJID,16,2)AS FL_CODE,A.DATAAREAID ")
                strSQL.Append(",C.IVZ_HOUSESIZEADJUSTMENTSQ_M,C.IVZ_HOUSESIZE,D.IVZ_AREAADJUSTMENTPRICE_SQWAH,C.IVZ_PROPERTYTYPEID,D.NAME,C.IVZ_TITLEDEEDNUM ")
                strSQL.Append(",E.DESCRIPTION AS PL_NAME,C.IVZ_HOUSENUM,A.REFPROPERTYSALESID,C.IVZ_LANDSIZEADJUSTMENTSQ_WAH,C.IVZ_LANDSIZE,B.IVZ_LandSizeAdjustmentSq_wah as LandSizePropertysales ")
                strSQL.Append(",B.IVZ_HOUSESIZEADJUSTMENTSQ_M as HouseSizePropertysales,A.CONTRACTNAME,B.NETAMOUNT,CT.NAME AS CUSTNAME ")
                strSQL.Append(",A.EXPECTTRANSFERDATE,DAY(A.EXPECTTRANSFERDATE) AS D_TRAN, MONTH(A.EXPECTTRANSFERDATE) AS M_TRAN, YEAR(A.EXPECTTRANSFERDATE) AS Y_TRAN ")
                strSQL.Append(",A.CONTRACTACCOUNT,b.DIMENSION3_,GETDATE()as NOW_DATE,YEAR(GETDATE()) as NOW_YEAR,MONTH(GETDATE()) as NOW_MONTH,DAY(GETDATE()) as NOW_DAY ")
                strSQL.Append(",isnull(case when C.IVZ_CADESTRALDATE = '' then 0 else YEAR(C.IVZ_CADESTRALDATE) end,0) AS depre_year, ")
                strSQL.Append("isnull(C.IVZ_HOUSESIZEADJUSTMENTSQ_M,0) AS HO_AREA,isnull(C.IVZ_LANDSIZEADJUSTMENTSQ_WAH,0) AS HO_AREA2, isnull(C.IVZ_CADESTRALAPPRAISALBUI40037,0) AS HO_PRICE ")
                strSQL.Append(",B.PROMOTIONAMOUNT,B.UNITPRICE,B.BASESALES, SUBSTRING(A.PROJID, 5, 1) as Projcode , SUBSTRING(C.ITEMID, 14,9) as PDCODE,isnull(GT.CashPromotionAmt,0) as CashPromotionAmt  ")
                strSQL.Append("FROM IVZ_PROPERTYSALESTABLE A ")
                strSQL.Append("LEFT JOIN IVZ_PROPERTYSALESLINE B ")
                strSQL.Append("ON A.PROPERTYSALESID = B.PROPERTYSALESID AND A.DATAAREAID = B.DATAAREAID ")
                strSQL.Append("LEFT JOIN INVENTTABLE C ")
                strSQL.Append("ON B.ITEMID = C.ITEMID ")
                strSQL.Append("LEFT JOIN PROJTABLE D ")
                strSQL.Append("ON A.PROJID = D.PROJID AND A.DATAAREAID = D.DATAAREAID ")
                strSQL.Append("LEFT OUTER JOIN IVZ_IMTYPE E ")
                strSQL.Append("ON C.IVZ_PROPERTYTYPEID = E.PROPERTYTYPEID AND SUBSTRING(A.PROJID,1,3) = E.PROJID AND A.DATAAREAID = E.DATAAREAID ")
                strSQL.Append("LEFT OUTER JOIN CUSTTABLE CT ")
                strSQL.Append("ON A.CONTRACTACCOUNT = CT.ACCOUNTNUM ")
                strSQL.Append("LEFT JOIN IVZ_PropertyPromotionTrans GT ")
                strSQL.Append("ON A.PROPERTYSALESID = GT.PROPERTYSALESID AND GT.ITEMID = 'PRO-0001-00009(Sale)' AND SELECTED = '1' ")
                'strSQL.Append("where a.PROPERTYSALESID = 'SO-SKV-AME1-000596' ")
                strSQL.Append("where a.PROPERTYSALESID = '" & paybank & "' ")
                strSQL.Append("AND A.PROPERTYSALESTATUS NOT IN ('0','6','7') ")
                dbERP.sql = strSQL.ToString
                Dim dt As DataTable = dbERP.QueryDataTable()
                If dbERP.RowsAffected >= 1 Then
                    projID = dt.Rows(0).Item("PROJID")
                    totalB = dt.Rows(0).Item("BASESALES")
                    Contract = dt.Rows(0).Item("CONTRACTACCOUNT")
                    proj = dt.Rows(0).Item("DIMENSION3_")
                    Now_Dates = dt.Rows(0).Item("NOW_DATE")
                    PD_Code = dt.Rows(0).Item("ITEMID")
                    BD_CODE = dt.Rows(0).Item("BD_CODE")
                    FL_CODE = dt.Rows(0).Item("FL_CODE")
                    m_home_area = dt.Rows(0).Item("HO_AREA")
                    m_home_price = dt.Rows(0).Item("HO_PRICE")
                    mTransYear = dt.Rows(0).Item("Y_TRAN")
                    PROPERTYSALESID = dt.Rows(0).Item("PROPERTYSALESID")
                    Projcode = dt.Rows(0).Item("Projcode")
                    landWAH = dt.Rows(0).Item("IVZ_LANDSIZEADJUSTMENTSQ_WAH")
                    NOW_DAY = dt.Rows(0).Item("NOW_DAY")
                    NOW_MONTH = dt.Rows(0).Item("NOW_MONTH")
                    NOW_YEAR = dt.Rows(0).Item("NOW_YEAR")
                    namepro = dt.Rows(0).Item("name")
                    CONTRACTNAME = dt.Rows(0).Item("CONTRACTNAME")
                    D_TRAN = dt.Rows(0).Item("D_TRAN")
                    M_TRAN = dt.Rows(0).Item("M_TRAN")
                    Y_TRAN = dt.Rows(0).Item("Y_TRAN")
                    PDCODE = dt.Rows(0).Item("PDCODE")
                    UNITPRICE = dt.Rows(0).Item("UNITPRICE")
                    PROMOTIONAMOUNT = dt.Rows(0).Item("PROMOTIONAMOUNT")
                    BASESALES = dt.Rows(0).Item("BASESALES")
                    '*********** ส่วนลดเงินสด ****************
                    CashPromotionAmt = dt.Rows(0).Item("CashPromotionAmt")

                    '*********** มูลค่าของโปรโมชั่น **********************
                    promotionlast = PROMOTIONAMOUNT - CashPromotionAmt

                    now_date = NOW_YEAR.ToString + "-" + NOW_MONTH.ToString + "-" + NOW_DAY.ToString

                    '******Get ปีที่เริ่มคิดค่าเสื่อม Year ,พื้นที่สิ่งปลูกสร้างล่าสุด, ราคาประเมินสิ่งปลูกสร้าง******
                    If dt.Rows(0).Item("depre_year") = 0 Then
                        m_depre = 0
                    Else
                        m_depre = dt.Rows(0).Item("depre_year")
                    End If
                    '**********************************************************************
                End If

                Dim totalC As Decimal
                totalC = totalB

                '********************* ชื่อบริษัท ********************'
                '*************************************************'
                Dim company As String
                Dim namecompany As String
                strSQL.Remove(0, strSQL.Length)
                strSQL.Append(" select A.IVZ_TAXNAME ,A.DATAAREAID from companyinfo A ")
                strSQL.Append("LEFT JOIN IVZ_PROPERTYSALESTABLE B ")
                strSQL.Append("ON A.DATAAREAID = B.DATAAREAID ")
                strSQL.Append("WHERE B.PROPERTYSALESID = '" & PROPERTYSALESID & "' ")
                strSQL.Append("AND B.PROPERTYSALESTATUS NOT IN ('0','6','7') ")
                dbERP.sql = strSQL.ToString
                Dim dtS As DataTable = dbERP.QueryDataTable()
                If dbERP.RowsAffected >= 1 Then
                    company = dtS.Rows(0).Item("DATAAREAID")
                    namecompany = dtS.Rows(0).Item("IVZ_TAXNAME")
                End If

                '************* เงื่อนไข เลขบัญชีบริษัท *************
                Dim NumberBank As String
                Dim BankName As String

                If company = "ary" Then
                    NumberBank = "00976-8"
                    BankName = "TBANK"
                ElseIf company = "amm" Then
                    NumberBank = "1227-4"
                    BankName = "TBANK"
                ElseIf company = "chs" Then
                    NumberBank = "1227-4"
                    BankName = "TBANK"
                ElseIf company = "whl" Then
                    NumberBank = "1245-2"
                    BankName = "TBANK"
                ElseIf company = "col" Then
                    NumberBank = "54015-7"
                    BankName = "SCB"
                Else
                End If

                '******************** ธนาคารการกู้ ********************'
                '***************************************************'
                Dim PREM_AMT As Decimal
                Dim DECORATE_AMT As Decimal
                Dim namebank As String = " "
                strSQL.Remove(0, strSQL.Length)
                strSQL.Append("select a.PJ_CODE,a.MA_RUNNO,a.PD_CODE,a.PAY_TYPE ")
                strSQL.Append(",b.BANK_CODE,b.BANK_BRANCH,b.LOAN_AMT ")
                strSQL.Append(",b.PREM_AMT,b.DECORATE_AMT, ")
                strSQL.Append("isnull(c.DESCRIPTION, '') as NAME ")
                strSQL.Append("from ARY_FIN_TRANS_TRANS a ")
                strSQL.Append("left join ARY_FIN_TRANS_TRANS_BANK b on a.MA_RUNNO = b.MA_RUNNO ")
                strSQL.Append("left join IVZ_PSLoanBank c on b.BANK_CODE = c.LOANBANK and b.DATAAREAID = c.DATAAREAID ")
                strSQL.Append("where a.MA_RUNNO = '" & PROPERTYSALESID & "' ")
                strSQL.Append("order by a.PJ_CODE,a.PD_CODE,b.BANK_SEQ ")
                dbERP.sql = strSQL.ToString
                Dim dtt As DataTable = dbERP.QueryDataTable()
                If dbERP.RowsAffected >= 1 Then
                    namebank = dtt.Rows(0).Item("NAME")
                    totalE = dtt.Rows(0).Item("LOAN_AMT")
                    PREM_AMT = dtt.Rows(0).Item("PREM_AMT")
                    DECORATE_AMT = dtt.Rows(0).Item("DECORATE_AMT")
                End If

                Dim totalF As Decimal
                totalF = totalE


                '******************** Propotysale ********************'
                '***************************************************'
                Dim totalA As Decimal = 0

                '******************** ยอดชำระ ********************'
                '***************************************************'
                strSQL.Remove(0, strSQL.Length)
                strSQL.Append(" with Settle as ( ")
                strSQL.Append("select ")
                strSQL.Append("ROW_NUMBER() over (PARTITION BY A.PROPERTYSALESID,A.PAYMENTTYPE,A.DOWNPERIOD,A.transid order by c.TRANSCOMPANY,c.TRANSRECID,c.ACCOUNTNUM )  AS rn ")
                strSQL.Append(",a.PROPERTYSALESID,a.PAYMENTTYPE,a.DOWNPERIOD,a.TOTALAMOUNT,c.OFFSETTRANSVOUCHER,b.DATAAREAID,b.RECID,b.ACCOUNTNUM,c.TRANSCOMPANY,c.TRANSRECID ")
                strSQL.Append("from IVZ_PSPAYMENTTRANS a ")
                strSQL.Append("left join CUSTTRANS b on (a.TRANSID = b.IVZ_PROJTRANSID and a.PROPERTYSALESID = b.IVZ_PROPERTYSALESID and a.PAYMENTTYPE = b.IVZ_PAYMENTTYPE and a.DOWNPERIOD = b.IVZ_DOWNPERIOD and a.INVOICEACCOUNT = b.ACCOUNTNUM) ")
                strSQL.Append("left join CUSTSETTLEMENT c on (b.DATAAREAID = c.TRANSCOMPANY and b.RECID = c.TRANSRECID and b.ACCOUNTNUM = c.ACCOUNTNUM) ")
                strSQL.Append("where a.PROPERTYSALESID = '" & PROPERTYSALESID & "' ")
                strSQL.Append("and a.PAYMENTTYPE in (1,2,3,5) ")
                strSQL.Append("and a.TRANSSTATUS != 0 ")
                strSQL.Append(") ")
                strSQL.Append("select  isnull(sum(E.SETTLEAMOUNTMST),0) as sumtotal ")
                strSQL.Append("from Settle D ")
                strSQL.Append("left join CUSTSETTLEMENT E on D.DATAAREAID = E.TRANSCOMPANY and D.RECID = E.TRANSRECID and D.ACCOUNTNUM = E.ACCOUNTNUM ")
                strSQL.Append("where rn = 1 ")
                dbERP.sql = strSQL.ToString
                Dim dtt2 As DataTable = dbERP.QueryDataTable()
                If dbERP.RowsAffected >= 1 Then
                    totalA = dtt2.Rows(0).Item("sumtotal")
                End If


                Dim c As Decimal = Nothing
                c = Decimal.Parse(totalA)
                Console.WriteLine(c)

                strSQL.Remove(0, strSQL.Length)
                strSQL.Append("select isnull(sum(SETTLEAMOUNTCUR - AMOUNTCUR),0) as sumAdvance ")
                strSQL.Append("from CustTrans ")
                strSQL.Append("where ACCOUNTNUM = '" & Contract & "' ")
                strSQL.Append("and TRANSTYPE = '15' ")
                strSQL.Append("and SETTLEAMOUNTCUR > AMOUNTCUR ")
                strSQL.Append("and DIMENSION3_ = '" & proj & "' ")
                strSQL.Append("and IVZ_PROPERTYSALESID = '" & PROPERTYSALESID & "' ")
                dbERP.sql = strSQL.ToString
                Dim dtt3 As DataTable = dbERP.QueryDataTable()
                If dtt3.Rows.Count >= 1 Then
                    '********** หัก ยอดเงินที่ชำระแล้ว ****************
                    totalA = totalA + dtt3.Rows(0).Item("sumAdvance")
                End If

                '********* ราคาขายสุทธิ **********
                Dim bases As Decimal
                bases = UNITPRICE - CashPromotionAmt

                '********** ยอดเหลือชำระ ณ วันโอน ****************
                Dim totalPayBank As Decimal
                totalPayBank = bases - totalA

                '********** ลูกค้าเหลือรับเงินกับธนาคาร ****************
                Dim CusAndBank As Decimal
                CusAndBank = totalF - totalPayBank


                '********** พื้นที่ ***********
                Dim m_area As Decimal = 0
                'Dim condo As Char = "H"
                'Dim projcode As String = 
                If Projcode = "H" Then
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                    strSQL.Append("WHERE PJ_CODE = '" & projID & "' ")
                    strSQL.Append("AND PD_CODE = '" & PD_Code & "' ")
                    strSQL.Append("AND BD_CODE = '" & BD_CODE & "' ")
                    strSQL.Append("AND FL_CODE = '" & FL_CODE & "' ")
                    dbERP.sql = strSQL.ToString
                    Dim BDcondo As DataTable = dbERP.QueryDataTable()
                    If BDcondo.Rows.Count >= 1 Then
                        m_area = BDcondo.Rows(0).Item("TOTAL_AREA")
                    End If

                Else
                    'm_area = dt.Rows(0).Item("IVZ_LANDSIZEADJUSTMENTSQ_WAH")
                    m_area = landWAH
                End If

                '***********คำนวนค่าต่างๆ*************
                '***************************************
                '***************************************
                If Projcode = "H" Then
                    '********** ค่าติดตั้งค่าไฟ***********
                    Dim Total2 As Decimal = 0
                    Dim Total2_2 As Decimal = 0
                    'Dim m_exp_amt As Decimal
                    Dim mDepre2 As Decimal
                    Dim ab As String
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '5' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db1 As DataTable = dbERP.QueryDataTable()
                    If db1.Rows.Count >= 1 Then

                        If db1.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total2 = db1.Rows(0).Item("EXP_AMT")
                            m_exp_amt = Total2

                        ElseIf db1.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total2 = db1.Rows(0).Item("EXP_PER_AREA")
                            Total2_2 = db1.Rows(0).Item("EXP_MONTH")
                            m_exp_amt = m_area * Total2 * Total2_2

                        ElseIf db1.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total2 = db1.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt = ((totalE + PREM_AMT + DECORATE_AMT) * Total2) / a

                        ElseIf db1.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total2 = db1.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total2) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total2) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total2) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total2) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total2) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total2) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total2) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total2) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total2) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total2) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total2) / a
                            If m_depre = 0 Then

                            Else
                                mDepre2 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre2) / 100)
                            End If

                            m_exp_amt = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If

                    End If



                    '********** ค่าส่วนกลาง ***********
                    Dim Total1 As Decimal = 0
                    Dim Total1_2 As Decimal = 0
                    'Dim m_exp_amt2 As Decimal
                    Dim mDepre As Decimal
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '1' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db2 As DataTable = dbERP.QueryDataTable()
                    If db2.Rows.Count >= 1 Then
                        If db2.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total1 = db2.Rows(0).Item("EXP_AMT")
                            m_exp_amt2 = Total1

                        ElseIf db2.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total1 = db2.Rows(0).Item("EXP_PER_AREA")
                            Total1_2 = db2.Rows(0).Item("EXP_MONTH")
                            m_exp_amt2 = m_area * Total1 * Total1_2

                        ElseIf db2.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total1 = db2.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt2 = ((totalE + PREM_AMT + DECORATE_AMT) * Total1) / a

                        ElseIf db2.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total1 = db2.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total1) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total1) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total1) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total1) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total1) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total1) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total1) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total1) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total1) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total1) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total1) / a
                            If m_depre = 0 Then

                            Else
                                mDepre = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre) / 100)
                            End If

                            m_exp_amt2 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If

                    End If


                    '********** ค่ารักษามาตฐานน้ำ ***********
                    Dim Total3 As Decimal = 0
                    Dim Total3_2 As Decimal = 0
                    'Dim m_exp_amt3 As Decimal
                    Dim mDepre3 As Decimal

                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '4' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db3 As DataTable = dbERP.QueryDataTable()
                    If db3.Rows.Count >= 1 Then
                        If db3.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_AMT")
                            m_exp_amt3 = Total3

                        ElseIf db3.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_PER_AREA")
                            Total3_2 = db3.Rows(0).Item("EXP_MONTH")
                            m_exp_amt3 = m_area * Total3 * Total3_2

                        ElseIf db3.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt3 = ((totalE + PREM_AMT + DECORATE_AMT) * Total3) / a

                        ElseIf db3.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total3) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total3) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total3) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total3) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total3) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total3) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total3) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total3) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total3) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total3) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total3) / a
                            If m_depre = 0 Then

                            Else
                                mDepre3 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre3) / 100)
                            End If

                            m_exp_amt3 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If

                    End If

                    '********** ค่าใช้จ่ายอื่นๆ ***********
                    Dim Total4 As Decimal = 0
                    Dim Total4_2 As Decimal = 0
                    'Dim m_exp_amt4 As Decimal
                    Dim mDepre4 As Decimal
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '9' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db4 As DataTable = dbERP.QueryDataTable()
                    If db4.Rows.Count >= 1 Then
                        If db4.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_AMT")
                            m_exp_amt4 = Total4

                        ElseIf db4.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_PER_AREA")
                            Total4_2 = db4.Rows(0).Item("EXP_MONTH")
                            m_exp_amt4 = m_area * Total4 * Total4_2

                        ElseIf db4.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt4 = ((totalE + PREM_AMT + DECORATE_AMT) * Total4) / a

                        ElseIf db4.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total4) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total4) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total4) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total4) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total4) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total4) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total4) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total4) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total4) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total4) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total4) / a
                            If m_depre = 0 Then

                            Else
                                mDepre4 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre4) / 100)
                            End If

                            m_exp_amt4 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If

                    End If

                    '********** ค่าใช้เบี้ยประกัน ***********
                    Dim Total5 As Decimal = 0
                    Dim Total5_2 As Decimal = 0
                    'Dim m_exp_amt5 As Decimal
                    Dim mDepre5 As Decimal
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '10' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db5 As DataTable = dbERP.QueryDataTable()
                    If db5.Rows.Count >= 1 Then
                        If db5.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total5 = db5.Rows(0).Item("EXP_AMT")
                            m_exp_amt5 = Total5

                        ElseIf db5.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total5 = db5.Rows(0).Item("EXP_PER_AREA")
                            Total5_2 = db5.Rows(0).Item("EXP_MONTH")
                            m_exp_amt5 = m_area * Total5 * Total5_2

                        ElseIf db5.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total5 = db5.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt5 = ((totalE + PREM_AMT + DECORATE_AMT) * Total5) / a

                        ElseIf db5.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total5 = db5.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total5) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total5) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total5) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total5) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total5) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total5) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total5) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total5) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total5) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total5) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total5) / a
                            If m_depre = 0 Then

                            Else
                                mDepre5 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre5) / 100)
                            End If

                            m_exp_amt5 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If

                    End If

                    '********** ค่ากองทุนอาคารชุด ***********
                    Dim Total6 As Decimal = 0
                    Dim Total6_2 As Decimal = 0
                    'Dim m_exp_amt6 As Decimal
                    Dim mDepre6 As Decimal
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '3' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db6 As DataTable = dbERP.QueryDataTable()
                    If db6.Rows.Count >= 1 Then
                        If db6.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total6 = db6.Rows(0).Item("EXP_AMT")
                            m_exp_amt6 = Total6

                        ElseIf db6.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total6 = db6.Rows(0).Item("EXP_PER_AREA")
                            Total6_2 = db6.Rows(0).Item("EXP_MONTH")
                            m_exp_amt6 = m_area * Total6 * Total6_2

                        ElseIf db6.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total6 = db6.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt6 = ((totalE + PREM_AMT + DECORATE_AMT) * Total6) / a

                        ElseIf db6.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total6 = db6.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total6) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total6) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total6) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total6) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total6) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total6) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total6) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total6) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total6) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total6) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total6) / a
                            If m_depre = 0 Then

                            Else
                                mDepre6 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre6) / 100)
                            End If

                            m_exp_amt6 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If
                    End If
                    '********** ค่ากองธรรมเนียมการจดจำนอง ***********
                    Dim Total7 As Decimal = 0
                    Dim Total7_2 As Decimal = 0
                    'Dim m_exp_amt7 As Decimal
                    Dim mDepre7 As Decimal
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '6' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db7 As DataTable = dbERP.QueryDataTable()
                    If db7.Rows.Count >= 1 Then
                        If db7.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total7 = db7.Rows(0).Item("EXP_AMT")
                            m_exp_amt7 = Total7

                        ElseIf db7.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total7 = db7.Rows(0).Item("EXP_PER_AREA")
                            Total7_2 = db7.Rows(0).Item("EXP_MONTH")
                            m_exp_amt7 = m_area * Total7 * Total7_2

                        ElseIf db7.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total7 = db7.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt7 = ((totalE + PREM_AMT + DECORATE_AMT) * Total7) / a

                        ElseIf db7.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total7 = db7.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total7) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total7) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total7) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total7) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total7) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total7) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total7) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total7) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total7) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total7) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total7) / a
                            If m_depre = 0 Then

                            Else
                                mDepre7 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre7) / 100)
                            End If

                            m_exp_amt7 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If
                    End If

                    '********** ค่ากองธรรมเนียมการโอนห้องชุด ***********
                    Dim Total8 As Decimal = 0
                    Dim Total8_2 As Decimal = 0
                    'Dim m_exp_amt8 As Decimal
                    Dim mDepre8 As Decimal
                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '6' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db8 As DataTable = dbERP.QueryDataTable()
                    If db8.Rows.Count >= 1 Then
                        If db8.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total8 = db8.Rows(0).Item("EXP_AMT")
                            m_exp_amt8 = Total8

                        ElseIf db8.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total8 = db8.Rows(0).Item("EXP_PER_AREA")
                            Total8_2 = db8.Rows(0).Item("EXP_MONTH")
                            m_exp_amt8 = m_area * Total8 * Total8_2

                        ElseIf db8.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total8 = db8.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt8 = ((totalE + PREM_AMT + DECORATE_AMT) * Total8) / a

                        ElseIf db8.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total8 = db8.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_CONDO ")
                            strSQL.Append("WHERE pj_code = '" & projID & "' ")
                            strSQL.Append("AND pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND bd_code = '" & BD_CODE & "' ")
                            strSQL.Append("AND fl_code = '" & FL_CODE & "' ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (dbESTM2.Rows(0).Item("LAND1_AREA") * dbESTM1.Rows(0).Item("LAND_PRICE") * Total8) / a
                            Dim mLand2 = (dbESTM2.Rows(0).Item("LAND2_AREA") * dbESTM1.Rows(0).Item("LAND2_PRICE") * Total8) / a
                            Dim mLand3 = (dbESTM2.Rows(0).Item("LAND3_AREA") * dbESTM1.Rows(0).Item("LAND3_PRICE") * Total8) / a
                            Dim mLand4 = (dbESTM2.Rows(0).Item("LAND4_AREA") * dbESTM1.Rows(0).Item("LAND4_PRICE") * Total8) / a
                            Dim mLand5 = (dbESTM2.Rows(0).Item("LAND5_AREA") * dbESTM1.Rows(0).Item("LAND5_PRICE") * Total8) / a
                            Dim mLand6 = (dbESTM2.Rows(0).Item("LAND6_AREA") * dbESTM1.Rows(0).Item("LAND6_PRICE") * Total8) / a
                            Dim mLand7 = (dbESTM2.Rows(0).Item("LAND7_AREA") * dbESTM1.Rows(0).Item("LAND7_PRICE") * Total8) / a
                            Dim mLand8 = (dbESTM2.Rows(0).Item("LAND8_AREA") * dbESTM1.Rows(0).Item("LAND8_PRICE") * Total8) / a
                            Dim mLand9 = (dbESTM2.Rows(0).Item("LAND9_AREA") * dbESTM1.Rows(0).Item("LAND9_PRICE") * Total8) / a
                            Dim mLand10 = (dbESTM2.Rows(0).Item("LAND10_AREA") * dbESTM1.Rows(0).Item("LAND10_PRICE") * Total8) / a

                            Dim mHouse = ((m_home_area * m_home_price) * Total8) / a
                            If m_depre = 0 Then

                            Else
                                mDepre8 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre8) / 100)
                            End If

                            m_exp_amt8 = (mLand + mLand2 + mLand3 + mLand4 + mLand5 + mLand6 + mLand7 + mLand8 + mLand9 + mLand10) + mHouse

                        End If
                    End If

                    '********* ผลรวมค่าต่างๆ ของ H ********

                    m_exp_amt_all = m_exp_amt + m_exp_amt2 + m_exp_amt3 + m_exp_amt4 + m_exp_amt5 + m_exp_amt6
                Else
                    '********** ค่าส่วนกลาง L ***********
                    Dim Total1 As Decimal = 0
                    Dim Total1_2 As Decimal = 0
                    'Dim m_exp_amt As Decimal
                    Dim mDepre1 As Decimal

                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '1' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db1 As DataTable = dbERP.QueryDataTable()
                    If db1.Rows.Count >= 1 Then

                        If db1.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total1 = db1.Rows(0).Item("EXP_AMT")
                            m_exp_amt = Total1

                        ElseIf db1.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total1 = db1.Rows(0).Item("EXP_PER_AREA")
                            Total1_2 = db1.Rows(0).Item("EXP_MONTH")
                            m_exp_amt = m_area * Total1 * Total1_2

                        ElseIf db1.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total1 = db1.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt = ((totalE + PREM_AMT + DECORATE_AMT) * Total1) / a

                        ElseIf db1.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total1 = db1.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append(" SELECT * FROM ARY_FIN_TRANS_PRICE a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (m_area * dbESTM1.Rows(0).Item("LAND_PRICE") * Total1) / a

                            Dim mHouse = (m_home_area * m_home_price) * Total1 / a

                            If m_depre = 0 Then

                            Else
                                mDepre1 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre1) / 100)
                            End If

                            m_exp_amt = mHouse + mLand

                        End If
                    End If
                    '********** ค่าติดตั้งไฟฟ้าและน้ำ L ***********
                    Dim Total2 As Decimal = 0
                    Dim Total2_2 As Decimal = 0
                    'Dim m_exp_amt2 As Decimal
                    Dim mDepre2 As Decimal

                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '2' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db2 As DataTable = dbERP.QueryDataTable()
                    If db2.Rows.Count >= 1 Then

                        If db2.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total2 = db2.Rows(0).Item("EXP_AMT")
                            m_exp_amt2 = Total2

                        ElseIf db2.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total2 = db2.Rows(0).Item("EXP_PER_AREA")
                            Total2_2 = db2.Rows(0).Item("EXP_MONTH")
                            m_exp_amt2 = m_area * Total2 * Total2_2

                        ElseIf db2.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total2 = db2.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt2 = ((totalE + PREM_AMT + DECORATE_AMT) * Total2) / a

                        ElseIf db2.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total2 = db2.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append(" SELECT * FROM ARY_FIN_TRANS_PRICE a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (m_area * dbESTM1.Rows(0).Item("LAND_PRICE") * Total2) / a

                            Dim mHouse = (m_home_area * m_home_price) * Total2 / a

                            If m_depre = 0 Then

                            Else
                                mDepre2 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre2) / 100)
                            End If

                            m_exp_amt2 = mHouse + mLand

                        End If
                    End If

                    '********** ค่าธรรมเนียมจดจำนอง L ***********
                    Dim Total3 As Decimal = 0
                    Dim Total3_2 As Decimal = 0
                    'Dim m_exp_amt3 As Decimal
                    Dim mDepre3 As Decimal

                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '6' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db3 As DataTable = dbERP.QueryDataTable()
                    If db3.Rows.Count >= 1 Then

                        If db3.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_AMT")
                            m_exp_amt3 = Total3

                        ElseIf db3.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_PER_AREA")
                            Total3_2 = db3.Rows(0).Item("EXP_MONTH")
                            m_exp_amt3 = m_area * Total3 * Total3_2

                        ElseIf db3.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt3 = ((totalE + PREM_AMT + DECORATE_AMT) * Total3) / a

                        ElseIf db3.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total3 = db3.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append(" SELECT * FROM ARY_FIN_TRANS_PRICE a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (m_area * dbESTM1.Rows(0).Item("LAND_PRICE") * Total3) / a

                            Dim mHouse = (m_home_area * m_home_price) * Total3 / a

                            If m_depre = 0 Then

                            Else
                                mDepre3 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre3) / 100)
                            End If

                            m_exp_amt3 = mHouse + mLand

                        End If
                    End If

                    '********** ค่าธรรมเนียมการโอนพร้อมสิ่งปลูกสร้าง L ***********
                    Dim Total4 As Decimal = 0
                    Dim Total4_2 As Decimal = 0
                    'Dim m_exp_amt4 As Decimal
                    Dim mDepre4 As Decimal

                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '7' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db4 As DataTable = dbERP.QueryDataTable()
                    If db4.Rows.Count >= 1 Then

                        If db4.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_AMT")
                            m_exp_amt4 = Total4

                        ElseIf db4.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_PER_AREA")
                            Total4_2 = db4.Rows(0).Item("EXP_MONTH")
                            m_exp_amt4 = m_area * Total4 * Total4_2

                        ElseIf db4.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt4 = ((totalE + PREM_AMT + DECORATE_AMT) * Total4) / a

                        ElseIf db4.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total4 = db4.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append(" SELECT * FROM ARY_FIN_TRANS_PRICE a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (m_area * dbESTM1.Rows(0).Item("LAND_PRICE") * Total4) / a

                            Dim mHouse = (m_home_area * m_home_price) * Total4 / a

                            If m_depre = 0 Then

                            Else
                                mDepre4 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre4) / 100)
                            End If

                            m_exp_amt4 = mHouse + mLand

                        End If
                    End If

                    '********** ค่าใช้จ่ายอื่นๆ L ***********
                    Dim Total5 As Decimal = 0
                    Dim Total5_2 As Decimal = 0
                    'Dim m_exp_amt5 As Decimal
                    Dim mDepre5 As Decimal

                    strSQL.Remove(0, strSQL.Length)
                    strSQL.Append("SELECT * FROM ARY_FIN_TRANS_EXP a ")
                    strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                    strSQL.Append("AND a.EXP_CODE    = '9' ")
                    strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                    strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                    strSQL.Append("FROM ARY_FIN_TRANS_EXP b ")
                    strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                    strSQL.Append("AND b.EXP_CODE    = a.EXP_CODE ")
                    strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                    dbERP.sql = strSQL.ToString
                    Dim db5 As DataTable = dbERP.QueryDataTable()
                    If db4.Rows.Count >= 1 Then

                        If db4.Rows(0).Item("EXP_AMT") <> 0 Then
                            Total5 = db4.Rows(0).Item("EXP_AMT")
                            m_exp_amt5 = Total5

                        ElseIf db4.Rows(0).Item("EXP_PER_AREA") <> 0 Then
                            Total5 = db4.Rows(0).Item("EXP_PER_AREA")
                            Total5_2 = db4.Rows(0).Item("EXP_MONTH")
                            m_exp_amt5 = m_area * Total5 * Total5_2

                        ElseIf db4.Rows(0).Item("EXP_RATE_LOAN") <> 0 Then
                            Total5 = db4.Rows(0).Item("EXP_RATE_LOAN")
                            m_exp_amt5 = ((totalE + PREM_AMT + DECORATE_AMT) * Total5) / a

                        ElseIf db4.Rows(0).Item("EXP_RATE_ESTM") <> 0 Then
                            Total5 = db4.Rows(0).Item("EXP_RATE_ESTM")

                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append(" SELECT * FROM ARY_FIN_TRANS_PRICE a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM1 As DataTable = dbERP.QueryDataTable()
                            If dbESTM1.Rows.Count >= 1 Then
                            End If
                            strSQL.Remove(0, strSQL.Length)
                            strSQL.Append("SELECT * FROM ARY_FIN_TRANS_PRICE_Fix a ")
                            strSQL.Append("WHERE a.pj_code     = '" & projID & "' ")
                            strSQL.Append("AND a.pd_code = '" & PD_Code & "' ")
                            strSQL.Append("AND a.start_date <= '" & now_date & "' ")
                            strSQL.Append("AND a.start_date  = (SELECT MAX(b.start_date) ")
                            strSQL.Append("FROM ARY_FIN_TRANS_PRICE_Fix b ")
                            strSQL.Append("WHERE(b.pj_code = a.pj_code) ")
                            strSQL.Append("AND b.pd_code = a.pd_code ")
                            strSQL.Append("AND b.start_date <= '" & now_date & "') ")
                            dbERP.sql = strSQL.ToString
                            Dim dbESTM2 As DataTable = dbERP.QueryDataTable()
                            If dbESTM2.Rows.Count >= 1 Then
                            End If

                            Dim mLand = (m_area * dbESTM1.Rows(0).Item("LAND_PRICE") * Total5) / a

                            Dim mHouse = (m_home_area * m_home_price) * Total5 / a

                            If m_depre = 0 Then

                            Else
                                mDepre5 = (mTransYear - m_depre) + b
                                mHouse = mHouse - ((mHouse * mDepre5) / 100)
                            End If

                            m_exp_amt5 = mHouse + mLand

                        End If
                    End If

                    '********** รวมค่าต่างๆ ***********

                    m_exp_amt_all = m_exp_amt + m_exp_amt2 + m_exp_amt3 + m_exp_amt4 + m_exp_amt5

                End If
                '********* บริษัทฯ ได้รับเงินจากลูกค้าเกินมา ********

                Dim totalall As Decimal
                totalall = CusAndBank - m_exp_amt_all


                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(1)

                If Projcode = "H" Then
                    'ชื่อโครงการแนวสูง
                    worksheet.Cells("W14").Value = "  " & namepro & " ห้อง " & PDCODE & "  "
                Else
                    'ชื่อโครงการแนวราบ
                    worksheet.Cells("W14").Value = "  " & namepro & " แปลง " & PDCODE & "  "
                End If
                'ชื่อลูกค้า
                worksheet.Cells("R15").Value = " " & CONTRACTNAME & ""
                'วันที่
                worksheet.Cells("W16").Value = " " & D_TRAN & "  " & Thai_Month_Array(M_TRAN) & " " & Y_TRAN + 543 & " "
                'ราคาขาย
                worksheet.Cells("AJ17").Value = " " & FormatNumber(UNITPRICE.ToString(), 2, TriState.True, ) & " "
                'หัก ส่วนลด
                worksheet.Cells("AJ18").Value = " " & FormatNumber(CashPromotionAmt.ToString(), 2, TriState.True, ) & " "
                'โปรโมชั่น
                worksheet.Cells("Z20").Value = " " & FormatNumber(promotionlast.ToString(), 2, TriState.True, ) & " "
                'ราคาขายสุทธิ
                'worksheet.Cells("AJ21").Value = " " & FormatNumber(bases.ToString(), 2, TriState.True, ) & " "
                'หักยอดที่ชำระแล้ว
                worksheet.Cells("AJ22").Value = " " & FormatNumber(totalA.ToString(), 2, TriState.True, ) & " "
                'ยอดเงินที่เหลือชำระ ณ วันโอน
                'worksheet.Cells("AJ23").Value = " " & FormatNumber(totalPayBank.ToString(), 2, TriState.True, ) & " "
                'ลูกค้ากู้ธนาคาร
                worksheet.Cells("O24").Value = " " & namebank & " "
                'ลูกค้ากู้ธนาคาร
                worksheet.Cells("AJ24").Value = " " & FormatNumber(totalF.ToString(), 2, TriState.True, ) & " "
                'ลูกค้าเหลือรับเงินกับธนาคาร
                'worksheet.Cells("AJ25").Value = " " & FormatNumber(CusAndBank.ToString(), 2, TriState.True, ) & " "
                'หักจ่ายค่าส่วนกลาง
                worksheet.Cells("AJ26").Value = " " & FormatNumber(m_exp_amt_all.ToString(), 2, TriState.True, ) & " "
                ''บริษัทได้รับเงินจากลูกค้าเกินมา
                'worksheet.Cells("G49").Value = " " & FormatNumber(totalall.ToString(), 2, TriState.True, ) & " "
                'โดย
                worksheet.Cells("F30").Value = " " & namebank & "  "
                'สั่งจ่ายกู้ส่วนเกินให้กับ
                worksheet.Cells("AB30").Value = " " & namecompany & " "
                'สั่งจ่ายให้กับ
                worksheet.Cells("J33").Value = " " & CONTRACTNAME & " "
                ''จำนวน
                'worksheet.Cells("H56").Value = " " & FormatNumber(totalall.ToString(), 2, TriState.True, ) & " บาท"
                'โดยสั่งจ่ายจากธนาคาร
                worksheet.Cells("P34").Value = " " & BankName & " "
                'บัญชีเลขที่
                worksheet.Cells("Z34").Value = " " & NumberBank & " "
                ''ลงวันที่
                'worksheet.Cells("L57").Value = " " & NOW_DAY & "  " & Thai_Month_Array(NOW_MONTH) & " " & NOW_YEAR + 543 & " "


                ' ** คิด % **
                Dim Perone As Decimal
                Dim Pertwo As Decimal
                Dim sumper As Decimal
                Perone = (CashPromotionAmt / UNITPRICE) * 100
                Pertwo = (promotionlast / UNITPRICE) * 100
                sumper = Perone + Pertwo
                Dim u As Decimal = 0
                Dim g As Decimal = BASESALES
                Dim sss As String
                u = UNITPRICE - PROMOTIONAMOUNT
                If g < u Or sumper = u Then
                    sss = " ▲ "
                Else
                    sss = " ▼ "
                End If
                'ส่วนลดโปรโมชั่น
                worksheet.Cells("AV37").Value = " " & FormatNumber(sumper.ToString(), 2, TriState.True, ) & " %  "
                worksheet.Cells("AY37").Value = " " & sss & " "

                worksheet.Cells("AV18").Value = "" & FormatNumber(Perone.ToString(), 2, TriState.True, ) & " % "

                worksheet.Cells("AV20").Value = "" & FormatNumber(Pertwo.ToString(), 2, TriState.True, ) & " % "

                ''ชื่อกรรมการ แนวราบหรือแนวสูง''
                If Projcode = "L" Then
                    worksheet.Cells("B3").Value = "" & namecompany & ""
                    worksheet.Cells("L5").Value = "" & NOW_DAY & "  " & Thai_Month_Array(NOW_MONTH + 1) & " " & NOW_YEAR + 543 & ""
                    worksheet.Cells("J6").Value = "คุณอาณัติ ปิ่นรัฒน์"
                    worksheet.Cells("J7").Value = "ผู้ช่วยกรรมการผู้จัดการอาวุโส"
                    worksheet.Cells("AJ6").Value = "คุณวิวัฒน์"
                    worksheet.Cells("AN40").Value = "(คุณอาณัติ ปิ่นรัฒน์)"
                    worksheet.Cells("AN41").Value = "ผู้ช่วยกรรมการผู้จัดการอาวุโส"


                Else
                    worksheet.Cells("B3").Value = "" & namecompany & ""
                    worksheet.Cells("L5").Value = "" & NOW_DAY & "  " & Thai_Month_Array(NOW_MONTH + 1) & " " & NOW_YEAR + 543 & ""
                    worksheet.Cells("J6").Value = "คุณวิศิษฎ์ เลาหพูนรังษี และ คุณนิภาพัฒน์"
                    worksheet.Cells("J7").Value = "โรมรัตนพันธ์ และ กรรมการผู้มีอำนาจลงนาม"
                    worksheet.Cells("AJ6").Value = "คุณวิศิษฎ์"
                    worksheet.Cells("AN40").Value = "(                                            )"
                    worksheet.Cells("AN41").Value = "กรรมการผู้มีอำนาจลงนาม"
                End If



                Dim file As String = "paybank2.xlsx"


                'strHtml = "<table id = 'tg-oabZt' border = '1' width=38%> "
                'strHtml += "<tr > "
                'strHtml += "<td  colspan='5' align = 'right'><font >ฝ่าย : <U>  การเงิน(การรับเงิน)</U></font><br>เอกสารเลขที่ : <U>  /2560</U> </td> "
                'strHtml += "</tr> "
                'strHtml += "<tr> <td align = 'center' colspan='5'><h3>บันทึกข้อความ</H3><H2> " & namecompany & " </H2> <br></td> "
                'strHtml += "</tr> "
                'strHtml += "<tr> "
                'strHtml += "<td><B>วัน/เดือน/ปี :</B><U> " & NOW_DAY & "  " & Thai_Month_Array(NOW_MONTH + 1) & " " & NOW_YEAR + 543 & "</U> "
                'strHtml += "<br><B>เรียน :</B><U> คุณอาณัติ ปิ่นรัตน์ <br> ผู้ช่วยกรรมการผู้จัดการอาวุโส</U> "
                'strHtml += "<br><B>เรื่อง :</B><U>ขออนุมัติจัดทำเช็คจ่ายคืนลูกค้า จำนวน 1 ฉบับ</U> "
                'strHtml += "<br><B>สำเนาถึง :</B><U> ฝ่ายบัญชี,การเงิน,การตลาดและฝ่ายขาย </U></td> "
                'strHtml += "<td><center><B><U>วัตถุประสงค์</U></B><BR><BR><input type='checkbox'>เพือโปรดทราบ<br><input type='checkbox'>เพื่อดำเนินการ<br><input type='checkbox' checked>เพื่อพิจารณาอนุมัติ<br><input type='checkbox'>เพื่อโปรดสั่งการ</center></td>"
                'strHtml += "<td colspan='3'><center><dd><B><U>ฝ่ายที่ต้องได้รับ</U></B><BR><BR><input type='checkbox'>คุณวิวัฒน์                 <input type='checkbox' checked>ฝ่ายบัญชี<br><input type='checkbox' checked>ฝ่ายการตลาด           <input type='checkbox'>ฝ่ายสารสนเทศ<br><input type='checkbox' checked>ฝ่ายขาย                   <input type='checkbox'>ฝ่ายลูกค้าสัมพันธ์<br><input type='checkbox' checked>ฝ่ายการเงิน              <input type='checkbox'>ฝ่ายบริการหลังการขาย</center></td>"
                'strHtml += "</tr> "
                'strHtml += "<tr > "
                'strHtml += "<td colspan='5'> <br>"
                'strHtml += " <dd>ขออนุมัติคืนเงินลูกค้า เนื่องจากบริษัทฯ ได้รับชำระเงิน ณ วันโอนกรรมสิทธิ์เกิน ซึ่งมีรายละเอียด ดังนี้ "
                'strHtml += "</td> "
                'strHtml += "</tr > "
                'strHtml += "<tr > "
                'strHtml += "<td colspan='5' style='border:0px'> <br>"
                'strHtml += " <dd>โครงการ/แปลง :   " & namepro & " ( " & PDCODE & " ) "
                'strHtml += "<BR>ชื่อลูกค้า :  " & CONTRACTNAME & ""
                'strHtml += "</td> "
                'strHtml += "</tr > "

                'strHtml += "<td colspan='2' style='border:0px'>"
                'strHtml += "<dd> "
                ''strHtml += "ชื่อลูกค้า : <br> "
                'strHtml += "วันที่โอน : <br> "
                'strHtml += "ราคาขาย : <br> "
                'strHtml += "(หัก) ส่วนลดเงินสด : <br> "
                'strHtml += "*มูลค่าโปรโมชั่นของแถมคิดเป็นจำนวน : xx,xxxx บาท   <br> "
                'strHtml += "ราคาขายสุทธิ :  <br> "
                'strHtml += "(หัก) ยอดเงินที่ชำระแล้ว : <br> "
                'strHtml += "ยอดเหลือชำระ ณ วันโอนฯ : <br> "
                'strHtml += "ลูกค้ากู้ " & namebank & " : <br> "
                'strHtml += "**ลูกค้าเหลือรับเงินกับธนาคาร : <BR> "
                'strHtml += "(1)หัก ธนาคารจ่ายค่าส่วนกลาง,<BR>ค่าธรรมเนียมจดจำนอง : <BR> "
                'strHtml += "บริษัทฯ ได้รับเงินจากลูกค้าเกินมา<br> "
                'strHtml += "</td> "
                'strHtml += "<td style='border:1px' colspan='2'> <br> "
                ''strHtml += " " & namepro & " <br> ( " & PDCODE & " )<br><br> "
                'strHtml += "  "

                'If M_TRAN = 0 Then
                '    strHtml += " <br> "
                'Else
                '    strHtml += " " & D_TRAN & "  " & Thai_Month_Array(M_TRAN - 1) & " " & Y_TRAN + 543 & " <br> "
                'End If

                'strHtml += " " & FormatNumber(UNITPRICE.ToString(), 2, TriState.True, ) & " บาท <br>  "
                'strHtml += " " & FormatNumber(CashPromotionAmt.ToString(), 2, TriState.True, ) & " บาท <br>  "
                'strHtml += " " & FormatNumber(promotionlast.ToString(), 2, TriState.True, ) & " บาท <br> "
                'strHtml += " " & FormatNumber(bases.ToString(), 2, TriState.True, ) & " บาท  <br> "
                'strHtml += " " & FormatNumber(totalA.ToString(), 2, TriState.True, ) & " บาท <br> "
                'strHtml += " " & FormatNumber(totalPayBank.ToString(), 2, TriState.True, ) & " บาท <br> "
                'strHtml += " " & FormatNumber(totalF.ToString(), 2, TriState.True, ) & " บาท <br> "
                'strHtml += " " & FormatNumber(CusAndBank.ToString(), 2, TriState.True, ) & " บาท .<BR><BR> "
                'strHtml += " " & FormatNumber(m_exp_amt_all.ToString(), 2, TriState.True, ) & " บาท .<BR> "
                'strHtml += " " & FormatNumber(totalall.ToString(), 2, TriState.True, ) & " บาท .<BR> "
                'strHtml += " <br> "
                'strHtml += "  "
                'strHtml += "</td> "
                'strHtml += "<td style='border:0px' colspan='1' >"

                '' ** คิด % **
                'Dim Perone As Decimal
                'Dim Pertwo As Decimal
                'Dim sumper As Decimal
                'Perone = (CashPromotionAmt / UNITPRICE) * 100
                'Pertwo = (promotionlast / UNITPRICE) * 100
                'sumper = Perone + Pertwo

                'strHtml += "<br><br>" & FormatNumber(Perone.ToString(), 2, TriState.True, ) & " %<br> " & FormatNumber(Pertwo.ToString(), 2, TriState.True, ) & " %<br><br><br><br><br><br><br><br><br>"
                'strHtml += "</td>"
                'strHtml += "</tr > "




                'strHtml += "<tr border = '9'> "
                'strHtml += "<td colspan='5' border = '9' > "
                'strHtml += "<br>โดย" & namebank & " ได้สั่งจ่ายเงินกู้ส่วนเกิน ให้แก่ " & namecompany & " และบริษัทได้รับไว้แล้ว ณ วันโอนฯ<br><br> "
                'strHtml += "<dd>ดังนั้น จึงขออนุมัติจัดทำเช็คเพื่อคืนเงินส่วนเกินให้แก่ลูกค้า ดังนี้<br> "
                'strHtml += "1.สั่งจ่าย " & CONTRACTNAME & " จำนวน " & FormatNumber(totalall.ToString(), 2, TriState.True, ) & " บาท<br> "
                'strHtml += "โดยสั่งจ่ายธนาคาร<U> &nbsp; &nbsp;" & BankName & " &nbsp; &nbsp;</U> บัญชีเลขที่<U> &nbsp; &nbsp;" & NumberBank & " &nbsp; &nbsp;</U> ลงวันที่ "
                'If M_TRAN = 0 Then
                '    strHtml += " <br><br> "
                'Else
                '    strHtml += "<U>" & NOW_DAY & "  " & Thai_Month_Array(NOW_MONTH + 1) & " " & NOW_YEAR + 543 & " </U> <br><br> "
                'End If
                'strHtml += "จึงเรียนมาเพื่อโปรดพิจรณาอนุมัติ<br></dd> <br><br>"
                'strHtml += "หมายเหตุ : <br> "
                'strHtml += "<br>"
                'strHtml += "</td > "

                'strHtml += "</tr> "
                'strHtml += "<tr>"
                'strHtml += "<td ALIGN = right colspan='5' > "
                'strHtml += "ส่วนลดเงินสด,โปรโมชั่น " & FormatNumber(sumper.ToString(), 2, TriState.True, ) & " %"
                'Dim u As Decimal = 0
                'Dim g As Decimal = BASESALES
                'u = UNITPRICE - PROMOTIONAMOUNT
                'If g < u Or sumper = u Then
                '    strHtml += "<b> ▲ </b>"
                'Else
                '    strHtml += "<b> ▼ </b>"
                'End If
                'strHtml += "</td> "
                'strHtml += "</tr>"
                ''strHtml += "<td style='border:0px'> </td> "
                'strHtml += "</tr> "
                'strHtml += "<tr > "
                'strHtml += "<td align = 'center' ><br>_______________<br>(นางสาวณัฐริกา จงใจงาม)<br>เจ้าหน้าที่การเงิน<br>___/___/_____</td> "
                'strHtml += "<td align = 'center' ><br>_______________<br>(คุณหนึ่งฤทัย เตจ๊ะ)<br>ผู้ช่วยผู้จัดการ ฝ่ายการเงิน<br>___/___/_____</td> "
                'strHtml += "<td align = 'center' ><br>_______________<br>(คุณพิระยา อาษา)<br>ผู้จัดการ ฝ่ายการเงิน<br>___/___/_____</td> "
                'strHtml += "<td align = 'center' colspan='2'><br>_______________<br>(คุณอาณัติ ปิ่นรัตน์)<br>ผู้ช่วยกรรมการผู้จัดการอาวุโส<br>___/___/_____</td> "
                'strHtml += "</tr> "
                'strHtml += "<tr > "
                'strHtml += "<td align = 'center' ><B>ผู้เสนอ</B></td> "
                'strHtml += "<td align = 'center' ><B>ผู้ตรวจสอบ</B></td> "
                'strHtml += "<td align = 'center' ><B>ผู้ควบคุม</B></td> "
                'strHtml += "<td align = 'center' colspan='2'><B>ผู้อนุมัติ</B></td> "

                'strHtml += "</tr>  "
                'strHtml += "</table> "

                Response.Clear() ' important'
                Response.AddHeader("Cache-Control", "max-age=0") ' optional'
                Response.ContentEncoding = Encoding.UTF8 ' optional'
                Response.HeaderEncoding = Encoding.UTF8 ' optional'
                Response.Charset = Encoding.UTF8.WebName ' optional'
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ' important'
                Response.AddHeader("content-disposition", "attachment;  filename=" & file) ' important'
                Response.BinaryWrite(package.GetAsByteArray()) ' important'
                Response.Flush() ' important - ensures content has been sent '
                Response.End() '
                'Response.Write(strHtml)
                'Response.End()

                '***********************' 
                '***** Create File *****' 
                '***********************' 
                'Dim mFileName As String
                'Dim mFileNameUrl As String
                'Dim datenows As Date = DateTime.Now
                'If strHtml.ToString <> "" Then
                '    mFileName = PD_Code.ToString
                '    mFileNameUrl = Server.MapPath(".") + "\text\" + mFileName + ".html"
                '    Dim fs As New FileStream(mFileNameUrl, FileMode.Create, FileAccess.Write)
                '    Dim s As New StreamWriter(fs)
                '    s.BaseStream.Seek(0, SeekOrigin.End)
                '    s.WriteLine(strHtml.ToString)
                '    s.Close()
                '    strHtml += "<br><div style='color: #FF00FF; background-color: #000000;'><center><a href='" + "open_excel_file.aspx?mFile=" + mFileName + "' target='_blank' style='color: #FF00FF;'><br><b>Download Excel File</b><br></a></center></div>"

                '    strHtml += "<b><br>รวมค่าใช้ค่าส่วนกลาง,ค่าเธรรมเนียมต่างๆ<br></b>"

                '    If Projcode = "L" Then
                '        strHtml += "ค่าส่วนกลาง " & FormatNumber(m_exp_amt.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่าติดตั้งไฟฟ้าและน้ำ " & FormatNumber(m_exp_amt2.ToString(), 2, TriState.True, ) & " บาท "
                '        strHtml += "<br>ค่าธรรมเนียมจดจำนอง " & FormatNumber(m_exp_amt3.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่าธรรมเนียมการโอนพร้อมสิ่งปลูกสร้าง " & FormatNumber(m_exp_amt4.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่าใช้จ่ายอื่นๆ " & FormatNumber(m_exp_amt5.ToString(), 2, TriState.True, ) & " บาท"
                '    Else
                '        strHtml += "ค่าติดตั้งค่าไฟ " & FormatNumber(m_exp_amt.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่าส่วนกลาง " & FormatNumber(m_exp_amt2.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่ารักษามาตฐานน้ำ " & FormatNumber(m_exp_amt3.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่าใช้จ่ายอื่นๆ  " & FormatNumber(m_exp_amt4.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่าใช้เบี้ยประกัน " & FormatNumber(m_exp_amt5.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่ากองทุนอาคารชุด " & FormatNumber(m_exp_amt6.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่ากองธรรมเนียมการจดจำนอง " & FormatNumber(m_exp_amt7.ToString(), 2, TriState.True, ) & " บาท"
                '        strHtml += "<br>ค่ากองธรรมเนียมการโอนห้องชุด " & FormatNumber(m_exp_amt8.ToString(), 2, TriState.True, ) & " บาท"

                '    End If



                'End If

                'Response.Write("<a href='#' style='background-color: #333333; color: #FF00FF;' onclick='hide_detail()'><b>ปิดรายละเอียด[X]</b></a>")


                'Response.Write(strHtml.ToString)
                'Response.End()
            End If
            End If
    End Sub


End Class
