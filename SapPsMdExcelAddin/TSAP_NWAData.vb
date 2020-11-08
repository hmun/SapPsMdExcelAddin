Public Class TSAP_NWAData

    Public aHdrRec As TDataRec
    Public aData As TData

    Private Hd_Fields() As String = {"I_NUMBER"}
    Private Data_Fields() As String = {"ACTIVITY", "CONTROL_KEY", "PLANT", "WORK_CNTR", "DESCRIPTION", "WBS_ELEMENT", "WORK_ACTIVITY", "UN_WORK", "UN_WORK_ISO", "CALCULATION_KEY", "ACTTYPE", "PRIORITY", "MILESTONE", "DISTRIBUTION_KEY", "NUMBER_OF_CAPACITIES", "PERCENT_OF_WORK", "EXEC_FACTOR", "STD_TEXT_KEY", "DURATION_NORMAL", "DURATION_NORMAL_UNIT", "DURATION_NORMAL_UNIT_ISO", "DURATION_MINIMUM", "DURATION_MINIMUM_UNIT", "DURATION_MINIMUM_UNIT_ISO", "FLEXIBLE_DURATION", "FACTORY_CALENDAR", "REDUCTION_STRATEGY", "INFO_REC", "VENDOR_NO", "PURCH_ORG", "PUR_GROUP", "MATL_GROUP", "AGREEMENT", "AGMT_ITEM", "PRICE", "PRICE_UNIT", "CURRENCY", "CURRENCY_ISO", "COST_ELEM", "PUR_INFO_RECORD_DATA_FIXED", "PLND_DELRY", "OPERATION_QTY", "OPERATION_MEASURE_UNIT", "OPERATION_MEASURE_UNIT_ISO", "NOT_MRP_APPLICABLE", "SORTED_BY", "PREQ_NAME", "GR_RCPT", "TRACKINGNO", "UNLOAD_PT", "CONSTRAINT_TYPE_START", "CONSTRAINT_TYPE_FINISH", "CONSTRAINT_START_DATE", "CONSTRAINT_START_TIME", "CONSTRAINT_FINISH_DATE", "CONSTRAINT_FINISH_TIME", "CONSTRAINT_TYPE_FINISH_FC", "CONSTRAINT_TYPE_START_FC", "CONSTRAINT_START_DATE_FC", "CONSTRAINT_START_TIME_FC", "CONSTRAINT_FINISH_DATE_FC", "CONSTRAINT_FINISH_TIME_FC", "BUS_AREA", "PROFIT_CTR", "CHANGE_NO", "REFERENCE_POINT", "CSTG_SHEET", "OVERHEAD_KEY", "TAXJURCODE", "OBJECTCLASS", "ACTIVITY_COSTS", "PROJECT_SUMMARIZATION", "REQ_COST_CENTER", "EARLY_START_DATE", "EARLY_START_TIME", "EARLY_FINISH_DATE", "EARLY_FINISH_TIME", "LATEST_START_DATE", "LATEST_START_TIME", "LATEST_FINISH_DATE", "LATEST_FINISH_TIME", "FLOAT_FREE", "FLOAT_TOTAL", "USER_FIELD_KEY", "USER_FIELD_CHAR20_1", "USER_FIELD_CHAR20_2", "USER_FIELD_CHAR10_1", "USER_FIELD_CHAR10_2", "USER_FIELD_QUAN1", "USER_FIELD_UNIT1", "USER_FIELD_UNIT1_ISO", "USER_FIELD_QUAN2", "USER_FIELD_UNIT2", "USER_FIELD_UNIT2_ISO", "USER_FIELD_CURR1", "USER_FIELD_CUKY1", "USER_FIELD_CUKY1_ISO", "USER_FIELD_CURR2", "USER_FIELD_CUKY2", "USER_FIELD_CUKY2_ISO", "USER_FIELD_DATE1", "USER_FIELD_DATE2", "USER_FIELD_FLAG1", "USER_FIELD_FLAG2", "DURATION_NORMAL_FC", "DURATION_NORMAL_FC_UNIT", "DURATION_NORMAL_FC_UNIT_ISO", "DURATION_MINIMUM_FC", "DURATION_MINIMUM_FC_UNIT", "DURATION_MINIMUM_FC_UNIT_ISO", "EARLY_START_DATE_FC", "EARLY_START_TIME_FC", "EARLY_FINISH_DATE_FC", "EARLY_FINISH_TIME_FC", "LATEST_START_DATE_FC", "LATEST_START_TIME_FC", "LATEST_FINISH_DATE_FC", "LATEST_FINISH_TIME_FC", "FLOAT_FREE_FC", "FLOAT_TOTAL_FC", "FUNC_AREA"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sWbs As String = "IT_ACTIVITY"

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        aHdrRec = New TDataRec(aIntPar)
        Dim aPostRec As New TDataRec(aIntPar)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        aPostRec = pData.getFirstRecord()
        If Not IsNothing(aPostRec) Then
            For Each aTStrRec In aPostRec.aTDataRecCol
                If valid_Hdr_Field(aTStrRec) Then
                    aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                End If
            Next
        End If
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        aData = New TData(aIntPar)
        fillData = True
        aCnt = 1
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            ' add the valid WBS fields
            For Each aTStrRec In aTDataRec.aTDataRecCol
                If valid_Data_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sWbs)
                End If
            Next
            aCnt += 1
        Next
    End Function

    Public Function valid_Hdr_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Hdr_Field = False
        If pTStrRec.Strucname = "" Or pTStrRec.Strucname = "HD" Then
            valid_Hdr_Field = isInArray(pTStrRec.Fieldname, Hd_Fields)
        End If
    End Function

    Public Function valid_Data_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Data_Field = False
        If pTStrRec.Strucname = "IT_ACTIVITY" Or pTStrRec.Strucname = "NWA" Then
            valid_Data_Field = isInArray(pTStrRec.Fieldname, Data_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Function getNetwork() As String
        Dim aTStrRec As SAPCommon.TStrRec
        getNetwork = ""
        For Each aTStrRec In aHdrRec.aTDataRecCol
            If aTStrRec.Fieldname = "I_NUMBER" Then
                getNetwork = aTStrRec.Value
                Exit Function
            End If
        Next
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("WBS_DBG", "DUMPHEADER") <> "", aIntPar.value("WBS_DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the WBS_DBG-DUMPHEADR Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
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

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("WBS_DBG", "DUMPDATA") <> "", aIntPar.value("WBS_DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the WBS_DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB As KeyValuePair(Of String, TDataRec)
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec(aIntPar)
            Dim aDataRec_Am As New TDataRec(aIntPar)
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB In aData.aTDataDic
                aDataRec = aKvB.Value
                Dim aFieldArray() As String = {}
                Dim aValueArray() As String = {}
                For Each aTStrRec In aDataRec.aTDataRecCol
                    Array.Resize(aFieldArray, aFieldArray.Length + 1)
                    aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                    Array.Resize(aValueArray, aValueArray.Length + 1)
                    aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                Next
                aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                aRange.Value = aFieldArray
                aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                aRange.Value = aValueArray
                i += 2
            Next
        End If
    End Sub

End Class
