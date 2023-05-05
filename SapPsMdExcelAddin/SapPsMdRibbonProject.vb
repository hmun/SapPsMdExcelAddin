' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapPsMdRibbonProject
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapPsMdRibbonProject getGenParametrs - " & "reading Parameter")
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

        Dim aSAPProjectDefinitionPI As New SAPProjectDefinitionPI(pSapCon, aIntPar)

        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("PROJ_WS", "DATA") <> "", aIntPar.value("PROJ_WS", "DATA"), "ProjectDefinition")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonProject.exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("PROJ_LOFF", "DATA") <> "", CInt(aIntPar.value("PROJ_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("PROJ_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("PROJ_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("PROJ_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("PROJ_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("PROJ_COL", "DATAMSG") <> "", aIntPar.value("PROJ_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("PROJ_RET", "OKMSG") <> "", aIntPar.value("PROJ_RET", "OKMSG"), "OK")
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
                ' Projects are handled line by line and not in packages. Using aItems is for standardization reasons - Will only contain one item.
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
                    Dim aTSAP_ProjectData
                    If pMode = "Change" Then
                        aTSAP_ProjectData = New TSAP_ProjectChgData(aPar, aIntPar)
                    Else
                        aTSAP_ProjectData = New TSAP_ProjectData(aPar, aIntPar)
                    End If
                    If aTSAP_ProjectData.fillProjectinfo(aItems) Then
                        ' check if we should dump this document
                        If aObjNr = aDumpObjNr Then
                            log.Debug("SapPsMdRibbonProject.exec - " & "dumping Object Nr " & CStr(aObjNr))
                            aTSAP_ProjectData.dumpProjectinfo()
                        End If
                        ' post the object here
                        If pMode = "Create" Then
                            log.Debug("SapPsMdRibbonProject.exec - " & "calling aSAPProjectDefinitionPI.createSingle")
                            aRetStr = aSAPProjectDefinitionPI.createSingle(aTSAP_ProjectData)
                            log.Debug("SapPsMdRibbonProject.exec - " & "aSAPProjectDefinitionPI.createSingle returned, aRetStr=" & aRetStr)
                            aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        ElseIf pMode = "Change" Then
                            log.Debug("SapPsMdRibbonProject.exec - " & "calling aSAPProjectDefinitionPI.changeSingle")
                            aRetStr = aSAPProjectDefinitionPI.changeSingle(aTSAP_ProjectData)
                            log.Debug("SapPsMdRibbonProject.exec - " & "aSAPProjectDefinitionPI.changeSingle returned, aRetStr=" & aRetStr)
                            aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        End If
                    Else
                        log.Warn("SapPsMdRibbonProject.exec - " & "filling Projectinfo in aTSAP_ProjectData failed!")
                        aDws.Cells(i, aMsgClmnNr) = "Filling Projectinfo in aTSAP_ProjectData failed!"
                    End If
                    aItems = New TData(aIntPar)
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonProject.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonProject.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonProject.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub Status(ByRef pSapCon As SapCon, Optional pMode As String = "Set")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPProjectDefinitionPI As New SAPProjectDefinitionPI(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aObjNr As UInt64 = 0
        Dim aProLOff As Integer = If(aIntPar.value("PROJ_LOFF", "DATA") <> "", CInt(aIntPar.value("PROJ_LOFF", "DATA")), 4)
        Dim aLOffCtrl As Integer = If(aIntPar.value("PROJ_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("PROJ_LOFFCTRL", "DATA")), 4)
        Dim aDumpObjNr As UInt64 = If(aIntPar.value("PROJ_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("PROJ_DBG", "DUMPOBJNR")), 0)
        Dim aProWsName As String = If(aIntPar.value("PROJ_WS", "DATA") <> "", aIntPar.value("PROJ_WS", "DATA"), "ProjectDefinition")
        Dim aProWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("PROJ_COL", "STATUSMSG") <> "", aIntPar.value("PROJ_COL", "STATUSMSG"), "INT-STATUSMSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("PROJ_RET", "OKMSG") <> "", aIntPar.value("PROJ_RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
        Try
            aProWs = aWB.Worksheets(aProWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aProWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        parseHeaderLine(aProWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=1)
        Try
            log.Debug("SapPsMdRibbonProject.SetStatus - " & "processing data - disabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPsMdExcelAddin.Application.EnableEvents = False
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aProLOff + 1
            Dim aKey As String
            Do
                aObjNr += 1
                If Left(CStr(aProWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aProItems As New TData(aIntPar)
                    '                    aProItems.addValue(aKey, CStr(aProWs.Cells(aProLOff - 3, 1).value), CStr(aProWs.Cells(i, 1).value), "", "")
                    Dim aTSAP_ProjStatusData As TSAP_ProjectStatusData
                    If pMode = "Get" Then
                        aTSAP_ProjStatusData = New TSAP_ProjectStatusData(aPar, aIntPar, aSAPProjectDefinitionPI, "GetStatus")
                    Else
                        aTSAP_ProjStatusData = New TSAP_ProjectStatusData(aPar, aIntPar, aSAPProjectDefinitionPI, "SetStatus")
                    End If
                    ' read DATA
                    aProItems.ws_parse_line_simple(aProWsName, aProLOff, i, jMax, pHdrLine:=1, pUplLine:=aLOffCtrl + 1)
                    If aTSAP_ProjStatusData.fillHeader(aProItems) Then
                        ' check if we should dump this document
                        If aObjNr = aDumpObjNr Then
                            log.Debug("SapPsMdRibbonProject.exec - " & "dumping Object Nr " & CStr(aObjNr))
                            aTSAP_ProjStatusData.dumpHeader()
                        End If
                        If pMode = "Get" Then
                            log.Debug("SapPsMdRibbonProject.GetStatus - " & "calling aSAPProjectDefinitionPI.GetStatus")
                            aRetStr = aSAPProjectDefinitionPI.GetStatus(aTSAP_ProjStatusData, pOKMsg:=aOKMsg)
                            log.Debug("SapPsMdRibbonProject.GetStatus - " & "aSAPProjectDefinitionPI.GetStatus returned, aRetStr=" & aRetStr)
                            aProWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                            ' output the data now
                            Dim aTData As TData
                            If aTSAP_ProjStatusData.aDataDic.aTDataDic.ContainsKey("E_SYSTEM_STATUS") Then
                                aTData = aTSAP_ProjStatusData.aDataDic.aTDataDic("E_SYSTEM_STATUS")
                                aTData.ws_output_line(aProWsName, "", i, jMax, pCoff:=0, pClear:=False, pKey:="")
                            End If
                            If aTSAP_ProjStatusData.aDataDic.aTDataDic.ContainsKey("E_USER_STATUS") Then
                                aTData = aTSAP_ProjStatusData.aDataDic.aTDataDic("E_USER_STATUS")
                                aTData.ws_output_line(aProWsName, "", i, jMax, pCoff:=0, pClear:=False, pKey:="")
                            End If
                        ElseIf pMode = "Set" Then
                            log.Debug("SapPsMdRibbonProject.SetStatus - " & "calling aSAPProjectDefinitionPI.SetStatus")
                            aRetStr = aSAPProjectDefinitionPI.SetStatus(aTSAP_ProjStatusData, pOKMsg:=aOKMsg)
                            log.Debug("SapPsMdRibbonProject.SetStatus - " & "aSAPProjectDefinitionPI.SetStatus returned, aRetStr=" & aRetStr)
                            aProWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aProWs.Cells(i, 1).value))
            log.Debug("SapPsMdRibbonProject.SetStatus - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPsMdExcelAddin.Application.EnableEvents = True
            Globals.SapPsMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapPsMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonProject.SetStatus failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md AddIn")
            log.Error("SapPsMdRibbonProject.SetStatus - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub
    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class
