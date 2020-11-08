Public Class TSAP_NetworkData
    Public aNetworkinfo As TDataRec

    Private Network_Fields() As String = {"NETWORK", "NETWORK_TYPE", "PROFILE", "PLANT", "MRP_CONTROLLER", "SHORT_TEXT", "START_DATE", "FINISH_DATE", "SCHED_TYPE", "START_DATE_FCST", "FINISH_DATE_FCST", "SCHED_TYPE_FCST", "NOT_AUTO_SCHEDULE", "NOT_AUTO_COSTING", "NOT_MRP_APPLICABLE", "PROJECT_DEFINITION", "WBS_ELEMENT", "SALES_DOC", "SALES_DOC_ITEM", "SUPERIOR_NETW", "SUPERIOR_NETW_ACT", "BUS_AREA", "PROFIT_CTR", "OBJECTCLASS", "TAXJURCODE", "PLANNER_GROUP", "CHANGE_NO", "PRIORITY", "EXEC_FACTOR", "COST_SHEET", "COST_VAR_PLAN", "COST_VAR_ACTUAL", "OVERHEAD_KEY", "SCHEDULING_EXACT_BREAK_TIMES", "NO_CAP_REQUIREMENTS", "CURRENCY", "FUNC_AREA"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sProj As String = "I_NETWORK"

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
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
                If valid_Network_Field(aTStrRec) Then
                    aNewNetworkinfo.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                End If
            Next
        End If
        aNetworkinfo = aNewNetworkinfo
        fillNetworkinfo = True
    End Function

    Public Function valid_Network_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Network_Field = False
        If pTStrRec.Strucname = "I_NETWORK" Then
            valid_Network_Field = isInArray(pTStrRec.Fieldname, Network_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
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
