﻿Public Class SapPsMdRibbonProject
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
                    Dim aTSAP_ProjectData As New TSAP_ProjectData(aPar, aIntPar)
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
                            ' log.Debug("SapPsMdRibbonProject.exec - " & "calling aSAPProjectDefinitionPI.changeSingle")
                            ' aRetStr = aSAPProjectDefinitionPI.changeSingle(aTSAP_CCData)
                            ' log.Debug("SapPsMdRibbonProject.exec - " & "aSAPProjectDefinitionPI.changeSingle returned, aRetStr=" & aRetStr)
                            ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
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

End Class
