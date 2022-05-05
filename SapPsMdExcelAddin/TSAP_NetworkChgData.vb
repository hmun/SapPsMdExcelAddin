' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TSAP_NetworkChgData
    Public aNetworkinfo As TDataRec

    Private Network_Fields_Chg() As String = {"MRP_CONTROLLER", "SHORT_TEXT", "START_DATE", "FINISH_DATE", "SCHED_TYPE", "START_DATE_FCST", "FINISH_DATE_FCST", "SCHED_TYPE_FCST", "NOT_AUTO_SCHEDULE", "NOT_AUTO_COSTING", "NOT_MRP_APPLICABLE", "PROJECT_DEFINITION", "WBS_ELEMENT", "SALES_DOC", "SALES_DOC_ITEM", "SUPERIOR_NETW", "SUPERIOR_NETW_ACT", "BUS_AREA", "PROFIT_CTR", "OBJECTCLASS", "TAXJURCODE", "PLANNER_GROUP", "CHANGE_NO", "PRIORITY", "EXEC_FACTOR", "COST_SHEET", "COST_VAR_PLAN", "COST_VAR_ACTUAL", "OVERHEAD_KEY", "SCHEDULING_EXACT_BREAK_TIMES", "NO_CAP_REQUIREMENTS", "CURRENCY", "FUNC_AREA"}
    Private Network_Fields_Upd() As String = {"MRP_CONTROLLER", "SHORT_TEXT", "START_DATE", "FINISH_DATE", "SCHED_TYPE", "START_DATE_FCST", "FINISH_DATE_FCST", "SCHED_TYPE_FCST", "NOT_AUTO_SCHEDULE", "NOT_AUTO_COSTING", "NOT_MRP_APPLICABLE", "PROJECT_DEFINITION", "WBS_ELEMENT", "SALES_DOC", "SALES_DOC_ITEM", "SUPERIOR_NETW", "SUPERIOR_NETW_ACT", "BUS_AREA", "PROFIT_CTR", "OBJECTCLASS", "TAXJURCODE", "PLANNER_GROUP", "CHANGE_NO", "PRIORITY", "EXEC_FACTOR", "COST_SHEET", "COST_VAR_PLAN", "COST_VAR_ACTUAL", "OVERHEAD_KEY", "SCHEDULING_EXACT_BREAK_TIMES", "NO_CAP_REQUIREMENTS", "CURRENCY", "FUNC_AREA"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sNetw_Chg As String = "I_NETWORK"
    Private Const sNetw_Upd As String = "I_NETWORK_UPD"

    Private aUseAsEmpty As String = "#"

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
        aUseAsEmpty = If(aIntPar.value("GEN", "USEASEMPTY") <> "", aIntPar.value("GEN", "USEASEMPTY"), "#")
    End Sub

    Public Function fillNetworkinfo(pData As TData) As Boolean
        aNetworkinfo = New TDataRec(aIntPar)
        Dim aFirstRec As New TDataRec(aIntPar)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewNetworkinfo As New TDataRec(aIntPar)
        Dim aNewExtinfo As New TDataRec(aIntPar)
        aFirstRec = pData.getFirstRecord()
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        If Not IsNothing(aFirstRec) Then
            For Each aTStrRec In aFirstRec.aTDataRecCol
                If valid_Network_Nr(aTStrRec) Then
                    aNewNetworkinfo.setValues("I_NUMBER", aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pUseAsEmpty:=aUseAsEmpty)
                ElseIf valid_Network_Field_Chg(aTStrRec) Then
                    aNewNetworkinfo.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pUseAsEmpty:=aUseAsEmpty)
                    If valid_Network_Field_Upd(aTStrRec) Then
                        aNewNetworkinfo.setValues(sNetw_Upd & "-" & aTStrRec.Fieldname, "X", "", "", pUseAsEmpty:=aUseAsEmpty)
                    End If
                End If
            Next
        End If
        aNetworkinfo = aNewNetworkinfo
        fillNetworkinfo = True
    End Function

    Public Function valid_Network_Nr(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Network_Nr = False
        If pTStrRec.Strucname = "I_NETWORK" And pTStrRec.Fieldname = "NETWORK" Then
            valid_Network_Nr = True
        End If
    End Function

    Public Function valid_Network_Field_Chg(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Network_Field_Chg = False
        If pTStrRec.Strucname = "I_NETWORK" Then
            valid_Network_Field_Chg = isInArray(pTStrRec.Fieldname, Network_Fields_Chg)
        End If
    End Function

    Public Function valid_Network_Field_Upd(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Network_Field_Upd = False
        If pTStrRec.Strucname = "I_NETWORK" Then
            valid_Network_Field_Upd = isInArray(pTStrRec.Fieldname, Network_Fields_Upd)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Function getNetwork() As String
        Dim aTStrRec As SAPCommon.TStrRec
        getNetwork = ""
        For Each aTStrRec In aNetworkinfo.aTDataRecCol
            If aTStrRec.Fieldname = "NETWORK" Then
                getNetwork = aTStrRec.Value
                Exit Function
            End If
        Next
    End Function

    Public Sub dumpNetworkinfo()
        Dim dumpHd As String = If(aIntPar.value("NETW_DBG", "DUMPDATA") <> "", aIntPar.value("NETW_DBG", "DUMPDATA"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpNetworkinfo - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the NETW_DBG-DUMPDATA Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
                Exit Sub
            End Try
            log.Debug("dumpNetworkinfo - " & "dumping to " & dumpHd)
            ' clear the Networkinfo
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Networkinfo
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aNetworkinfo.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

End Class
