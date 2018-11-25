Attribute VB_Name = "KdCommon"
Option Explicit

'�������������
Public Function KDRound(ByVal NumValue As Variant, Optional ByVal RScale As Long = 2) As Variant
    KDRound = CCur(Format(NumValue, "#,##0." & String$(RScale, "0")))
End Function

Public Function KDRoundDec(ByVal txtValue As String, Optional ByVal RScale As Long = 2) As Variant
    KDRoundDec = KDRound(Val(Replace(txtValue, ",", "")), RScale)
End Function

'����ָ���ı��ֺ����ڻ�ȡ��Ӧ�Ĺ�˾����
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

''���ݵ������ں͸�����������������
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
''���ݵ������ں͸�����������������
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

'���Լ�������һ���汾
''classtypeid ֻ����1007736 ����Ӧ��ϵͳ�ĸ�������
''1007737��Ӧ��ϵͳ���տ�����
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
'    '���岢����������
'    Dim StartDay As Date
'    If ds("FFstStDate").Value = 0 Then  '��ʾ����������
'        StartDay = billDate
'    Else '=1��ʾ��������ĩ����
'        Dim lastday As Date
'        lastday = DateAdd("m", 1, billDate)
'        lastday = DateAdd("d", -1, CDate(Year(lastday) & "-" & Month(lastday) & "-01"))
'        StartDay = lastday
'    End If
'
'    Dim PreDate As Date '����Ԥ�տ���
'    If ds("FOptMode").Value = 0 Then  '0��ʾ��������������
'        CalCreditDateEx = DateAdd("d", ds("FDay").Value, StartDay)
'        Exit Function
'    Else '1��ʾ���½᷽ʽ����
'        If ds("FLstDay").Value = 0 Then  '����
'            PreDate = DateAdd("m", ds("FDayMon").Value, StartDay)
'            CalCreditDateEx = CDate(Year(PreDate) & "-" & Month(PreDate) & "-" & CStr(ds("FDate").Value))
'            Exit Function
'        Else '����
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


