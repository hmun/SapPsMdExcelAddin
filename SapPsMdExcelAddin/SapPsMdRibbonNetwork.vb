Public Class SapPsMdRibbonNetwork
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapPsMdRibbonNetwork getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPPsMd"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function
    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub exec(ByRef pSapCon As SapCon, Optional pMode As String = "Create")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPNetworkPI As New SAPNetworkPI(pSapCon)
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("NETW_WS", "DATA") <> "", aIntPar.value("NETW_WS", "DATA"), "Network")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Network Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonNetwork.exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("NETW_LOFF", "DATA") <> "", CInt(aIntPar.value("NETW_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("NETW_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("NETW_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("NETW_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("NETW_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("NETW_COL", "DATAMSG") <> "", aIntPar.value("NETW_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("NETW_RET", "OKMSG") <> "", aIntPar.value("NETW_RET", "OKMSG"), "OK")
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPsMdExcelAddin.Application.EnableEvents = False
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 4, jMax + 1).value) <> ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' Networks are handled line by line and not in packages. Using aItems is for standardization reasons - Will only contain one item.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn And
                            CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "" And CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "N" Then
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 4, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(aLOff - 2, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    ' aItem = aItems.aTDataDic(aKey)
                    Dim aTSAP_NetworkData As New TSAP_NetworkData(aPar, aIntPar)
                    If aTSAP_NetworkData.fillNetworkinfo(aItems) Then
                        ' check if we should dump this document
                        If aObjNr = aDumpObjNr Then
                            log.Debug("SapPsMdRibbonNetwork.exec - " & "dumping Object Nr " & CStr(aObjNr))
                            aTSAP_NetworkData.dumpNetworkinfo()
                        End If
                        ' post the object here
                        If pMode = "Create" Then
                            log.Debug("SapPsMdRibbonNetwork.exec - " & "calling aSAPNetworkPI.createSingle")
                            aRetStr = aSAPNetworkPI.createFromData(aTSAP_NetworkData)
                            log.Debug("SapPsMdRibbonNetwork.exec - " & "aSAPNetworkPI.createSingle returned, aRetStr=" & aRetStr)
                            aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        ElseIf pMode = "Change" Then
                            ' log.Debug("SapPsMdRibbonNetwork.exec - " & "calling aSAPNetworkPI.change")
                            ' aRetStr = aSAPNetworkPI.change(aTSAP_CCData)
                            ' log.Debug("SapPsMdRibbonNetwork.exec - " & "aSAPNetworkPI.change returned, aRetStr=" & aRetStr)
                            ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        End If
                    Else
                        log.Warn("SapPsMdRibbonNetwork.exec - " & "filling Networkinfo in aTSAP_NetworkData failed!")
                        aDws.Cells(i, aMsgClmnNr) = "Filling Networkinfo in aTSAP_NetworkData failed!"
                    End If
                    aItems = New TData(aIntPar)
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonNetwork.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonNetwork.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonNetwork.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub execNWA(ByRef pSapCon As SapCon, Optional pMode As String = "Create")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPNetworkPI As New SAPNetworkPI(pSapCon)
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("NWA_WS", "DATA") <> "", aIntPar.value("NWA_WS", "DATA"), "NetworkActivity")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Network Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonNetwork.execNWA - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("NWA_LOFF", "DATA") <> "", CInt(aIntPar.value("NWA_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("NWA_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("NWA_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("NWA_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("NWA_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("NWA_COL", "DATAMSG") <> "", aIntPar.value("NWA_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aNetwClmn As String = If(aIntPar.value("NWA_COL", "NETW") <> "", aIntPar.value("NWA_COL", "NETW"), "I_NUMBER")
            Dim aNetwClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("NWA_RET", "OKMSG") <> "", aIntPar.value("NWA_RET", "OKMSG"), "Success")
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPsMdExcelAddin.Application.EnableEvents = False
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                ElseIf CStr(aDws.Cells(1, jMax).value) = aNetwClmn Then
                    aNetwClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 4, jMax + 1).value) <> ""
            Dim aNetwork As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' NetworksActivities are handled in packages per Network.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn And
                            CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "" And CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "N" Then
                            aNetwork = CStr(aDws.Cells(i, aNetwClmnNr).value)
                            aKey = aNetwork & "-" & CStr(i)
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 4, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(aLOff - 2, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    If aNetwork <> CStr(aDws.Cells(i + 1, aNetwClmnNr).value) Then
                        Dim aTSAP_NWAData As New TSAP_NWAData(aPar, aIntPar)
                        If aTSAP_NWAData.fillHeader(aItems) And aTSAP_NWAData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapPsMdRibbonNetwork.execNWA - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_NWAData.dumpHeader()
                                aTSAP_NWAData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Then
                                log.Debug("SapPsMdRibbonNetwork.execNWA - " & "calling aSAPNetworkPI.actCreateMultiple")
                                aRetStr = aSAPNetworkPI.actCreateMultiple(aTSAP_NWAData)
                                log.Debug("SapPsMdRibbonNetwork.execNWA - " & "aSAPNetworkPI.actCreateMultiple returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    Dim aKeyPair() As String
                                    aKeyPair = Split(aKey, "-")
                                    aDws.Cells(CInt(aKeyPair(1)), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                ' log.Debug("SapPsMdRibbonNetwork.execNWA - " & "calling aSAPNetworkPI.actChangeMultiple")
                                ' aRetStr = aSAPNetworkPI.actChangeMultiple(aTSAP_CCData)
                                ' log.Debug("SapPsMdRibbonNetwork.execNWA - " & "aSAPNetworkPI.actChangeMultiple returned, aRetStr=" & aRetStr)
                                ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                            End If
                        Else
                            log.Warn("SapPsMdRibbonNetwork.execNWA - " & "Filling Header or Data in aTSAP_NWAData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_NWAData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonNetwork.execNWA - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonNetwork.execNWA failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonNetwork.execNWA - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub execNWAE(ByRef pSapCon As SapCon, Optional pMode As String = "Create")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPNetworkPI As New SAPNetworkPI(pSapCon)
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("NWAE_WS", "DATA") <> "", aIntPar.value("NWAE_WS", "DATA"), "NetwActElement")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Network Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("NWAE_LOFF", "DATA") <> "", CInt(aIntPar.value("NWAE_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("NWAE_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("NWAE_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("NWAE_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("NWAE_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("NWAE_COL", "DATAMSG") <> "", aIntPar.value("NWAE_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aNetwClmn As String = If(aIntPar.value("NWAE_COL", "NETW") <> "", aIntPar.value("NWAE_COL", "NETW"), "I_NUMBER")
            Dim aNetwClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("NWAE_RET", "OKMSG") <> "", aIntPar.value("NWAE_RET", "OKMSG"), "Success")
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPsMdExcelAddin.Application.EnableEvents = False
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                ElseIf CStr(aDws.Cells(1, jMax).value) = aNetwClmn Then
                    aNetwClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 4, jMax + 1).value) <> ""
            Dim aNetwork As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' NetworksActivityElements are handled in packages per Network.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn And
                            CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "" And CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "N" Then
                            aNetwork = CStr(aDws.Cells(i, aNetwClmnNr).value)
                            aKey = aNetwork & "-" & CStr(i)
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 4, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(aLOff - 2, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    If aNetwork <> CStr(aDws.Cells(i + 1, aNetwClmnNr).value) Then
                        Dim aTSAP_NWAEData As New TSAP_NWAEData(aPar, aIntPar)
                        If aTSAP_NWAEData.fillHeader(aItems) And aTSAP_NWAEData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_NWAEData.dumpHeader()
                                aTSAP_NWAEData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Then
                                log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "calling aSAPNetworkPI.actCreateMultiple")
                                aRetStr = aSAPNetworkPI.actElemCreateMultiple(aTSAP_NWAEData)
                                log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "aSAPNetworkPI.actCreateMultiple returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    Dim aKeyPair() As String
                                    aKeyPair = Split(aKey, "-")
                                    aDws.Cells(CInt(aKeyPair(1)), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                ' log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "calling aSAPNetworkPI.actChangeMultiple")
                                ' aRetStr = aSAPNetworkPI.actChangeMultiple(aTSAP_CCData)
                                ' log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "aSAPNetworkPI.actChangeMultiple returned, aRetStr=" & aRetStr)
                                ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                            End If
                        Else
                            log.Warn("SapPsMdRibbonNetwork.execNWAE - " & "Filling Header or Data in aTSAP_NWAEData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_NWAEData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonNetwork.execNWAE - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonNetwork.execNWAE failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonNetwork.execNWAE - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub execCOMP(ByRef pSapCon As SapCon, Optional pMode As String = "Create")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPNetworkPI As New SAPNetworkPI(pSapCon)
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("COMP_WS", "DATA") <> "", aIntPar.value("COMP_WS", "DATA"), "Component")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Network Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("COMP_LOFF", "DATA") <> "", CInt(aIntPar.value("COMP_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("COMP_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("COMP_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("COMP_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("COMP_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("COMP_COL", "DATAMSG") <> "", aIntPar.value("COMP_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aNetwClmn As String = If(aIntPar.value("COMP_COL", "NETW") <> "", aIntPar.value("COMP_COL", "NETW"), "I_NUMBER")
            Dim aNetwClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("COMP_RET", "OKMSG") <> "", aIntPar.value("COMP_RET", "OKMSG"), "Success")
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPsMdExcelAddin.Application.EnableEvents = False
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                ElseIf CStr(aDws.Cells(1, jMax).value) = aNetwClmn Then
                    aNetwClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 4, jMax + 1).value) <> ""
            Dim aNetwork As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' Networks Components are handled in packages per Network.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn And
                            CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "" And CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "N" Then
                            aNetwork = CStr(aDws.Cells(i, aNetwClmnNr).value)
                            aKey = aNetwork & "-" & CStr(i)
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 4, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(aLOff - 2, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    If aNetwork <> CStr(aDws.Cells(i + 1, aNetwClmnNr).value) Then
                        Dim aTSAP_CompData As New TSAP_CompData(aPar, aIntPar)
                        If aTSAP_CompData.fillHeader(aItems) And aTSAP_CompData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_CompData.dumpHeader()
                                aTSAP_CompData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Then
                                log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "calling aSAPNetworkPI.addComponent")
                                aRetStr = aSAPNetworkPI.addComponent(aTSAP_CompData)
                                log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "aSAPNetworkPI.addComponent returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    Dim aKeyPair() As String
                                    aKeyPair = Split(aKey, "-")
                                    aDws.Cells(CInt(aKeyPair(1)), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                ' log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "calling aSAPNetworkPI.changeComponent")
                                ' aRetStr = aSAPNetworkPI.changeComponent(aTSAP_CompData)
                                ' log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "aSAPNetworkPI.changeComponent returned, aRetStr=" & aRetStr)
                                ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                            End If
                        Else
                            log.Warn("SapPsMdRibbonNetwork.execCOMP - " & "Filling Header or Data in aTSAP_CompData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_CompData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonNetwork.execCOMP - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonNetwork.execCOMP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonNetwork.execCOMP - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

End Class
