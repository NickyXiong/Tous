Attribute VB_Name = "KdCommon"
Option Explicit

'金蝶的四舍五入
Public Function KDRound(ByVal NumValue As Variant, Optional ByVal RScale As Long = 2) As Variant
    KDRound = CCur(Format(NumValue, "#,##0." & String$(RScale, "0")))
End Function

Public Function KDRoundDec(ByVal txtValue As String, Optional ByVal RScale As Long = 2) As Variant
    KDRoundDec = KDRound(Val(Replace(txtValue, ",", "")), RScale)
End Function

'根据指定的币种和日期获取相应的公司汇率
Public Function GetExchange(lCurrency As Long, dDate As Date) As Variant
    Dim ecb As New EBCGL.CRateType
    Dim dct As KFO.Dictionary
    Set dct = ecb.GetRateByCurrency(1, dDate, lCurrency)
    Set ecb = Nothing
    If dct.Lookup("FExchangeRate") Then
        GetExchange = dct("FExchangeRate")
    Else
        GetExchange = 0
    End If
End Function

Public Function GetCurrencyOperator(lCurrency As Long) As String
    Dim ecb As New EBCGL.CurrencySet
    Dim cur As EBCGL.Currencyx
    If ecb.ExistenceCheck(lCurrency) = 0 Then
        GetCurrencyOperator = "*"
    Else
        Set cur = ecb.item(lCurrency)
        GetCurrencyOperator = cur.Operator
    End If
    Set ecb = Nothing
End Function

''根据单据日期和付款条件，计算账期
'Public Function GetCreditDate(payNumber As String, billDate) As Date
'    Dim ebgl As New Payset
'    Dim payid As Long
'    payid = ebgl.ExistenceCheck(, payNumber)
'    If payid = 0 Then
'        Exit Function
'    Else
'        Dim dic As New KFO.Dictionary
'        Dim vec As New KFO.Vector
'        dic("FCustID") = 0
'        dic("FEmpId") = 0
'        dic("FDeptID") = 0
'        dic("FBillDate") = billDate
'        dic("FID") = payid
'        vec.Add dic
'
'        Dim obj As Object
'        Set obj = CreateObject("K3MSaleCredit.clsGetData")
'        Call obj.GetSettleDate(MMTS.PropsString, vec)
'
'        GetCreditDate = vec("FSettleDate")
'        Set obj = Nothing
'    End If
'    Set ebgl = Nothing
'End Function

'
''根据单据日期和付款条件，计算账期
'Public Function CalCreditDate(payid As Long, billDate) As Date
'    Dim dic As New KFO.Dictionary
'    Dim vec As New KFO.Vector
'    dic("FCustID") = 0
'    dic("FEmpId") = 0
'    dic("FDeptID") = 0
'    dic("FBillDate") = billDate
'    dic("FID") = payid
'    vec.Add dic
'
'    Dim obj As Object
'    Set obj = CreateObject("K3MSaleCredit.clsGetData")
'    Call obj.GetSettleDate(MMTS.PropsString, dic)
'
'    CalCreditDate = dic("FSettleDate")
'    Set obj = Nothing
'End Function

'是自己开发的一个版本
''classtypeid 只能是1007736 用于应付系统的付款条件
''1007737用应收系统的收款条件
'Public Function CalCreditDateEx(payid As Long, billDate As Date) As Date
'    Dim sqlcomm As New SqlAdapter
'
'    Dim ds As ADODB.Recordset
'    Set ds = sqlcomm.ExeSqlString("select * from t_payColConditionEntry Where FID=" & payid)
'    If ds.BOF And ds.EOF Then
'        CalCreditDateEx = billDate
'        Exit Function
'    End If
'    Set sqlcomm = Nothing
'
'    ds.MoveFirst
'    '定义并计算起算日
'    Dim StartDay As Date
'    If ds("FFstStDate").Value = 0 Then  '表示按单据日期
'        StartDay = billDate
'    Else '=1表示按单据月末日期
'        Dim lastday As Date
'        lastday = DateAdd("m", 1, billDate)
'        lastday = DateAdd("d", -1, CDate(Year(lastday) & "-" & Month(lastday) & "-01"))
'        StartDay = lastday
'    End If
'
'    Dim PreDate As Date '定义预收款日
'    If ds("FOptMode").Value = 0 Then  '0表示按信用天数结算
'        CalCreditDateEx = DateAdd("d", ds("FDay").Value, StartDay)
'        Exit Function
'    Else '1表示按月结方式结算
'        If ds("FLstDay").Value = 0 Then  '按月
'            PreDate = DateAdd("m", ds("FDayMon").Value, StartDay)
'            CalCreditDateEx = CDate(Year(PreDate) & "-" & Month(PreDate) & "-" & CStr(ds("FDate").Value))
'            Exit Function
'        Else '按天
'            PreDate = DateAdd("d", ds("FDayMon").Value, StartDay)
'            Dim SettleDay As Date
'            SettleDay = CDate(Year(PreDate) & "-" & Month(PreDate) & "-" & CStr(ds("FDate").Value))
'            If PreDate <= SettleDay Then
'                CalCreditDateEx = SettleDay
'                Exit Function
'            Else
'                CalCreditDateEx = DateAdd("m", 1, SettleDay)
'                Exit Function
'            End If
'        End If
'    End If
'End Function


Public Function SafeSqlString(s As String) As String
    SafeSqlString = Replace(s, "'", "''")
End Function


