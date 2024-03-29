﻿' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPNetworkPI

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
        End Try
    End Sub

    Private Sub addToFieldArray(pArrayName As String, pFieldName As String, ByRef pFieldsDic As Dictionary(Of String, String()))
        Dim aArray As String()
        If pFieldsDic.ContainsKey(pArrayName) Then
            aArray = pFieldsDic(pArrayName)
            Array.Resize(aArray, aArray.Length + 1)
            aArray(aArray.Length - 1) = pFieldName
            pFieldsDic.Remove(pArrayName)
            pFieldsDic.Add(pArrayName, aArray)
        Else
            aArray = {pFieldName}
            pFieldsDic.Add(pArrayName, aArray)
        End If
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_SetStatus(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {"NUMBER", "UNDO_SYSTEM_STATUS", "UNDO_USER_STATUS", "SET_SYSTEM_STATUS", "SET_USER_STATUS"}
        Dim aTables As String() = {"I_ACTIVITY_SYSTEM_STATUS", "I_ACTIVITY_USER_STATUS"}
        Try
            log.Debug("getMeta_SetStatus - " & "creating Function BAPI_BUS2002_SET_STATUS")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_SET_STATUS")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_SetStatus - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_GetStatus(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {"NUMBER", "WITHOUT_ACTIVITIES"}
        Dim aTables As String() = {"E_SYSTEM_STATUS", "E_USER_STATUS", "E_ACTIVITY_SYSTEM_STATUS", "E_ACTIVITY_USER_STATUS"}
        Try
            log.Debug("getMeta_GetStatus - " & "creating Function BAPI_BUS2002_GET_STATUS")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_GET_STATUS")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_GetStatus - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function createFromData(pData As TSAP_NetworkData, Optional pOKMsg As String = "OK") As String
        createFromData = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_CREATE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            oRETURN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aNetworkinfo.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                createFromData = createFromData & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    createFromData = createFromData & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            createFromData = If(createFromData = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & createFromData, "Error" & createFromData))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            createFromData = "Error: Exception in createFromData"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function change(pData As TSAP_NetworkChgData, Optional pOKMsg As String = "OK") As String
        change = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_CHANGE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            oRETURN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aNetworkinfo.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                change = change & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    change = change & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            change = If(change = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & change, "Error" & change))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            change = "Error: Exception in change"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function actCreateMultiple(pData As TSAP_NWAData, Optional pOKMsg As String = "OK") As String
        actCreateMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_ACT_CREATE_MULTI")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_ACTIVITY As IRfcTable = oRfcFunction.GetTable("IT_ACTIVITY")
            oRETURN.Clear()
            oIT_ACTIVITY.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_ACTIVITYAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_ACTIVITY"
                            If Not oIT_ACTIVITYAppended Then
                                oIT_ACTIVITY.Append()
                                oIT_ACTIVITYAppended = True
                            End If
                            oIT_ACTIVITY.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                actCreateMultiple = actCreateMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    actCreateMultiple = actCreateMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            actCreateMultiple = If(actCreateMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & actCreateMultiple, "Error" & actCreateMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            actCreateMultiple = "Error: Exception in actCreateMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function actChangeMultiple(pData As TSAP_NWAData, Optional pOKMsg As String = "OK") As String
        actChangeMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_ACT_CHANGE_MULTI")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_ACTIVITY As IRfcTable = oRfcFunction.GetTable("IT_ACTIVITY")
            Dim oIT_UPDATE_ACTIVITY As IRfcTable = oRfcFunction.GetTable("IT_UPDATE_ACTIVITY")
            oRETURN.Clear()
            oIT_ACTIVITY.Clear()
            oIT_UPDATE_ACTIVITY.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_ACTIVITYAppended As Boolean = False
                Dim oIT_UPDATE_ACTIVITYAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_ACTIVITY"
                            If Not oIT_ACTIVITYAppended Then
                                oIT_ACTIVITY.Append()
                                oIT_ACTIVITYAppended = True
                            End If
                            oIT_ACTIVITY.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "IT_UPDATE_ACTIVITY"
                            If Not oIT_UPDATE_ACTIVITYAppended Then
                                oIT_UPDATE_ACTIVITY.Append()
                                oIT_UPDATE_ACTIVITYAppended = True
                            End If
                            oIT_UPDATE_ACTIVITY.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                actChangeMultiple = actChangeMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    actChangeMultiple = actChangeMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            actChangeMultiple = If(actChangeMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & actChangeMultiple, "Error" & actChangeMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            actChangeMultiple = "Error: Exception in actChangeMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function actElemCreateMultiple(pData As TSAP_NWAEData, Optional pOKMsg As String = "OK") As String
        actElemCreateMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_ACTELEM_CREATE_M")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_ACT_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_ACT_ELEMENT")
            oRETURN.Clear()
            oIT_ACT_ELEMENT.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_ACT_ELEMENTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_ACT_ELEMENT"
                            If Not oIT_ACT_ELEMENTAppended Then
                                oIT_ACT_ELEMENT.Append()
                                oIT_ACT_ELEMENTAppended = True
                            End If
                            oIT_ACT_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                actElemCreateMultiple = actElemCreateMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    actElemCreateMultiple = actElemCreateMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            actElemCreateMultiple = If(actElemCreateMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & actElemCreateMultiple, "Error" & actElemCreateMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            actElemCreateMultiple = "Error: Exception in actElemCreateMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function actElemChangeMultiple(pData As TSAP_NWAEData, Optional pOKMsg As String = "OK") As String
        actElemChangeMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_ACTELEM_CHANGE_M")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_ACT_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_ACT_ELEMENT")
            Dim oIT_UPDATE_ACT_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_UPDATE_ACT_ELEMENT")
            oRETURN.Clear()
            oIT_ACT_ELEMENT.Clear()
            oIT_UPDATE_ACT_ELEMENT.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_ACT_ELEMENTAppended As Boolean = False
                Dim oIT_UPDATE_ACT_ELEMENTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_ACT_ELEMENT"
                            If Not oIT_ACT_ELEMENTAppended Then
                                oIT_ACT_ELEMENT.Append()
                                oIT_ACT_ELEMENTAppended = True
                            End If
                            oIT_ACT_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "IT_UPDATE_ACT_ELEMENT"
                            If Not oIT_UPDATE_ACT_ELEMENTAppended Then
                                oIT_UPDATE_ACT_ELEMENT.Append()
                                oIT_UPDATE_ACT_ELEMENTAppended = True
                            End If
                            oIT_UPDATE_ACT_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                actElemChangeMultiple = actElemChangeMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    actElemChangeMultiple = actElemChangeMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            actElemChangeMultiple = If(actElemChangeMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & actElemChangeMultiple, "Error" & actElemChangeMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            actElemChangeMultiple = "Error: Exception in actElemChangeMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function addComponent(pData As TSAP_CompData, Optional pOKMsg As String = "OK") As String
        addComponent = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_NETWORK_COMP_ADD")
            RfcSessionManager.BeginContext(destination)
            Dim oE_MESSAGE_TABLE As IRfcTable = oRfcFunction.GetTable("E_MESSAGE_TABLE")
            Dim oRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim oI_COMPONENTS_ADD As IRfcTable = oRfcFunction.GetTable("I_COMPONENTS_ADD")
            oE_MESSAGE_TABLE.Clear()
            oI_COMPONENTS_ADD.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oI_COMPONENTS_ADDAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "I_COMPONENTS_ADD"
                            If Not oI_COMPONENTS_ADDAppended Then
                                oI_COMPONENTS_ADD.Append()
                                oI_COMPONENTS_ADDAppended = True
                            End If
                            oI_COMPONENTS_ADD.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            If oI_COMPONENTS_ADD.Count > 0 Then
                oRfcFunction.Invoke(destination)
                Dim aErr As Boolean = False
                If oRETURN.GetValue("TYPE") = "E" Then
                    aErr = True
                End If
                For i As Integer = 0 To oE_MESSAGE_TABLE.Count - 1
                    addComponent = addComponent & ";" & oE_MESSAGE_TABLE(i).GetValue("MESSAGE_TEXT")
                Next i
                If aErr = False Then
                    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                    aSAPBapiTranctionCommit.commit(pWait:="X")
                End If
                addComponent = If(addComponent = "", pOKMsg, If(aErr = False, pOKMsg & addComponent, "Error" & addComponent))
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            addComponent = "Error: Exception in addComponent"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function changeComponent(pData As TSAP_CompData, Optional pOKMsg As String = "OK") As String
        changeComponent = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_NETWORK_COMP_CHANGE")
            RfcSessionManager.BeginContext(destination)
            Dim oE_MESSAGE_TABLE As IRfcTable = oRfcFunction.GetTable("E_MESSAGE_TABLE")
            Dim oRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim oI_COMPONENTS_CHANGE As IRfcTable = oRfcFunction.GetTable("I_COMPONENTS_CHANGE")
            Dim oI_COMPONENTS_CHANGE_UPDATE As IRfcTable = oRfcFunction.GetTable("I_COMPONENTS_CHANGE_UPDATE")
            oE_MESSAGE_TABLE.Clear()
            oI_COMPONENTS_CHANGE.Clear()
            oI_COMPONENTS_CHANGE_UPDATE.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oI_COMPONENTS_CHANGEAppended As Boolean = False
                Dim oI_COMPONENTS_CHANGE_UPDATEAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "I_COMPONENTS_CHANGE"
                            If Not oI_COMPONENTS_CHANGEAppended Then
                                oI_COMPONENTS_CHANGE.Append()
                                oI_COMPONENTS_CHANGEAppended = True
                            End If
                            oI_COMPONENTS_CHANGE.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "I_COMPONENTS_CHANGE_UPDATE"
                            If Not oI_COMPONENTS_CHANGE_UPDATEAppended Then
                                oI_COMPONENTS_CHANGE_UPDATE.Append()
                                oI_COMPONENTS_CHANGE_UPDATEAppended = True
                            End If
                            oI_COMPONENTS_CHANGE_UPDATE.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' get the Component Numbers
            Dim aCompDic As Dictionary(Of String, String)
            Dim aKey As String
            Dim aSAPFormat As New SAPCommon.SAPFormat()
            If oI_COMPONENTS_CHANGE.Count > 0 Then
                getListComponent(oRfcFunction.GetValue("NUMBER"), aCompDic)
                For i As Integer = 0 To oI_COMPONENTS_CHANGE.Count - 1
                    aKey = oRfcFunction.GetValue("NUMBER") & "-" & oI_COMPONENTS_CHANGE(i).GetValue("ACTIVITY") & "-" & oI_COMPONENTS_CHANGE(i).GetValue("ITEM_NUMBER")
                    If aCompDic.ContainsKey(aKey) Then
                        oI_COMPONENTS_CHANGE(i).SetValue("COMPONENT", aCompDic(aKey))
                        oI_COMPONENTS_CHANGE_UPDATE(i).SetValue("COMPONENT", aCompDic(aKey))
                    Else
                        changeComponent = changeComponent & ";" & "Component Number not found for Key: " & aKey
                        oI_COMPONENTS_CHANGE.Delete(i)
                    End If
                Next i
            End If
            ' call the BAPI
            If oI_COMPONENTS_CHANGE.Count > 0 Then
                oRfcFunction.Invoke(destination)
                Dim aErr As Boolean = False
                If oRETURN.GetValue("TYPE") = "E" Then
                    aErr = True
                End If
                For i As Integer = 0 To oE_MESSAGE_TABLE.Count - 1
                    changeComponent = changeComponent & ";" & oE_MESSAGE_TABLE(i).GetValue("MESSAGE_TEXT")
                Next i
                If aErr = False Then
                    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                    aSAPBapiTranctionCommit.commit(pWait:="X")
                End If
                changeComponent = If(changeComponent = "", pOKMsg, If(aErr = False, pOKMsg & changeComponent, "Error" & changeComponent))
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            changeComponent = "Error: Exception in changeComponent"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function getListComponent(pNumber As String, ByRef pCompDict As Dictionary(Of String, String), Optional pActivity As String = "", Optional pOKMsg As String = "OK") As String
        getListComponent = ""
        Dim aRfcFunction As IRfcFunction
        Try
            aRfcFunction = destination.Repository.CreateFunction("BAPI_NETWORK_COMP_GETDETAIL")
            RfcSessionManager.BeginContext(destination)
            Dim oE_COMPONENTS_DETAIL As IRfcTable = aRfcFunction.GetTable("E_COMPONENTS_DETAIL")
            Dim oRETURN As IRfcStructure = aRfcFunction.GetStructure("RETURN")
            Dim oI_ACTIVITY_RANGE As IRfcTable = aRfcFunction.GetTable("I_ACTIVITY_RANGE")
            oE_COMPONENTS_DETAIL.Clear()
            oI_ACTIVITY_RANGE.Clear()
            aRfcFunction.SetValue("NUMBER", pNumber)
            If pActivity <> "" Then
                oI_ACTIVITY_RANGE.Append()
                oI_ACTIVITY_RANGE.SetValue("SIGN", "I")
                oI_ACTIVITY_RANGE.SetValue("OPTION", "EQ")
                oI_ACTIVITY_RANGE.SetValue("LOW", pActivity)
            End If
            ' call the BAPI
            aRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            If oRETURN.GetValue("TYPE") = "E" Then
                aErr = True
            End If
            If oRETURN.GetValue("TYPE") <> "" Then
                getListComponent = getListComponent & ";" & oRETURN.GetValue("MESSAGE_TEXT")
            End If
            getListComponent = If(getListComponent = "", pOKMsg, If(aErr = False, pOKMsg & getListComponent, "Error" & getListComponent))
            Dim aKey As String
            pCompDict = New Dictionary(Of String, String)
            If aErr = False Then
                For i As Integer = 0 To oE_COMPONENTS_DETAIL.Count - 1
                    aKey = oE_COMPONENTS_DETAIL(i).GetValue("NETWORK") & "-" & oE_COMPONENTS_DETAIL(i).GetValue("ACTIVITY") & "-" & oE_COMPONENTS_DETAIL(i).GetValue("ITEM_NUMBER")
                    pCompDict.Add(aKey, oE_COMPONENTS_DETAIL(i).GetValue("COMPONENT"))
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            getListComponent = "Error: Exception in getListComponent"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function SetStatus(pData As TSAP_NetworkStatusData, Optional pOKMsg As String = "OK") As String
        SetStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_SET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oI_ACTIVITY_SYSTEM_STATUS As IRfcTable = oRfcFunction.GetTable("I_ACTIVITY_SYSTEM_STATUS")
            Dim oI_ACTIVITY_USER_STATUS As IRfcTable = oRfcFunction.GetTable("I_ACTIVITY_USER_STATUS")
            Dim oE_RESULT As IRfcTable = oRfcFunction.GetTable("E_RESULT")
            oI_ACTIVITY_SYSTEM_STATUS.Clear()
            oI_ACTIVITY_USER_STATUS.Clear()
            oE_RESULT.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            pData.aDataDic.to_IRfcTable(pKey:="I_ACTIVITY_SYSTEM_STATUS", pIRfcTable:=oI_ACTIVITY_SYSTEM_STATUS)
            pData.aDataDic.to_IRfcTable(pKey:="I_ACTIVITY_USER_STATUS", pIRfcTable:=oI_ACTIVITY_USER_STATUS)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            SetStatus = SetStatus & ";" & sRETURN.GetValue("MESSAGE")
            If sRETURN.GetValue("TYPE") = "E" Then
                aErr = True
            End If
            For i As Integer = 0 To oE_RESULT.Count - 1
                SetStatus = SetStatus & ";" & oE_RESULT(i).GetValue("STATUS_ACTION") & "-" & oE_RESULT(i).GetValue("STATUS_TYPE") & "-" & oE_RESULT(i).GetValue("MESSAGE_TEXT")
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    SetStatus = SetStatus & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            SetStatus = If(SetStatus = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & SetStatus, "Error" & SetStatus))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            SetStatus = "Error: Exception in SetStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function GetStatus(pData As TSAP_NetworkStatusData, pWITHOUT_ACTIVITIES As String, Optional pOKMsg As String = "OK") As String
        GetStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2002_GET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oE_SYSTEM_STATUS As IRfcTable = oRfcFunction.GetTable("E_SYSTEM_STATUS")
            Dim oE_USER_STATUS As IRfcTable = oRfcFunction.GetTable("E_USER_STATUS")
            Dim oE_ACTIVITY_SYSTEM_STATUS As IRfcTable = oRfcFunction.GetTable("E_ACTIVITY_SYSTEM_STATUS")
            Dim oE_ACTIVITY_USER_STATUS As IRfcTable = oRfcFunction.GetTable("E_ACTIVITY_USER_STATUS")
            oE_SYSTEM_STATUS.Clear()
            oE_USER_STATUS.Clear()
            oE_ACTIVITY_SYSTEM_STATUS.Clear()
            oE_ACTIVITY_USER_STATUS.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            oRfcFunction.SetValue("WITHOUT_ACTIVITIES", pWITHOUT_ACTIVITIES)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            GetStatus = GetStatus & ";" & sRETURN.GetValue("MESSAGE")
            If sRETURN.GetValue("TYPE") = "E" Then
                aErr = True
            End If
            If aErr = False Then
                ' return the system status
                pData.aDataDic.addValues(oTable:=oE_SYSTEM_STATUS, pStrucName:="E_SYSTEM_STATUS")
                ' return the user status
                pData.aDataDic.addValues(oTable:=oE_USER_STATUS, pStrucName:="E_USER_STATUS")
                Dim aCount As Integer
                Dim i As Integer
                If pData.aActivity <> "" Then
                    aCount = oE_ACTIVITY_SYSTEM_STATUS.Count - 1
                    i = 0
                    Do While i <= aCount
                        If oE_ACTIVITY_SYSTEM_STATUS(i).GetValue("ACTIVITY") <> pData.aActivity Then
                            oE_ACTIVITY_SYSTEM_STATUS.Delete(i)
                            aCount -= 1
                        Else
                            i += 1
                        End If
                    Loop
                    aCount = oE_ACTIVITY_USER_STATUS.Count - 1
                    i = 0
                    Do While i <= aCount
                        If oE_ACTIVITY_USER_STATUS(i).GetValue("ACTIVITY") <> pData.aActivity Then
                            oE_ACTIVITY_USER_STATUS.Delete(i)
                            aCount -= 1
                        Else
                            i += 1
                        End If
                    Loop
                End If
                ' return the activity system status
                pData.aDataDic.addValues(oTable:=oE_ACTIVITY_SYSTEM_STATUS, pStrucName:="E_ACTIVITY_SYSTEM_STATUS")
                ' return the activity user status
                pData.aDataDic.addValues(oTable:=oE_ACTIVITY_USER_STATUS, pStrucName:="E_ACTIVITY_USER_STATUS")
            End If
            GetStatus = If(GetStatus = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & GetStatus, "Error" & GetStatus))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNetworkPI")
            GetStatus = "Error: Exception in GetStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class


