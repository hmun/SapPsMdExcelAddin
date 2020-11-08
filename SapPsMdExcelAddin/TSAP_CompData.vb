Public Class TSAP_CompData

    Public aHdrRec As TDataRec
    Public aData As TData

    Private Hd_Fields() As String = {"NUMBER"}
    Private Data_Fields() As String = {"ACTIVITY", "TYPE_OF_PUR_RESV", "ITEM_NUMBER", "MATERIAL", "PLANT", "ENTRY_QUANTITY", "BASE_UOM", "BASE_UOM_ISO", "ITEM_CAT", "ITEM_TEXT", "MRP_RELEVANT", "REQ_DATE", "MANUAL_REQUIREMENTS_DATE", "LEAD_TIME_OFFSET_OPR", "LEAD_TIME_OFFSET_OPR_UNIT", "LEAD_TIME_OFFSET_OPR_UNIT_ISO", "MRP_DISTRIBUTION_KEY", "COST_RELEVANT", "STGE_LOC", "BATCH", "BOMEXPL_NO", "DELIVERY_DAYS", "PUR_GROUP", "PURCH_ORG", "INFO_REC", "PRICE", "PRICE_UNIT", "CURRENCY", "CURRENCY_ISO", "PUR_INFO_RECORD_DATA_FIXED", "AGREEMENT", "AGMT_ITEM", "GL_ACCOUNT", "VENDOR_NO", "GR_PR_TIME", "MATL_GROUP", "PREQ_NAME", "GR_RCPT", "TRACKINGNO", "UNLOAD_PT", "SORT_STRING", "BACKFLUSH", "BULK_MAT", "VSI_SIZE1", "VSI_SIZE2", "VSI_SIZE3", "VSI_SIZE_UNIT", "VSI_SIZE_UNIT_ISO", "VSI_QTY", "VAR_SIZE_COMP_MEASURE_UNIT", "VAR_SIZE_COMP_MEASURE_UNIT_ISO", "VSI_FORMULA", "VSI_NO", "ORIGINAL_QUANTITY", "ADDR_NO", "ADDR_NO2", "SUPP_VENDOR", "CUSTOMER", "WBS_ELEMENT", "S_ORD_ITEM", "MATERIAL_EXTERNAL", "MATERIAL_GUID", "MATERIAL_VERSION", "MATERIAL_LONG"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sWbs As String = "I_COMPONENTS_ADD"

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
        If pTStrRec.Strucname = "I_COMPONENTS_ADD" Or pTStrRec.Strucname = "COMP" Then
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
        Dim dumpHd As String = If(aIntPar.value("COMP_DBG", "DUMPHEADER") <> "", aIntPar.value("COMP_DBG", "DUMPHEADER"), "")
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
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the COMP_DBG-DUMPHEADR Parameter",
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
        Dim dumpDt As String = If(aIntPar.value("COMP_DBG", "DUMPDATA") <> "", aIntPar.value("COMP_DBG", "DUMPDATA"), "")
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
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the COMP_DBG-DUMPDATA Parameter",
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
