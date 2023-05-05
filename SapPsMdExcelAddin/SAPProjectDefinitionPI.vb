' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPProjectDefinitionPI
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr

    Sub New(aSapCon As SapCon, ByRef pIntPar As SAPCommon.TStr)
        aIntPar = pIntPar
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
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
        Dim aImports As String() = {"PROJECT_DEFINITION", "UNDO_SYSTEM_STATUS", "UNDO_USER_STATUS", "SET_SYSTEM_STATUS", "SET_USER_STATUS"}
        Dim aTables As String() = {}
        Try
            log.Debug("getMeta_SetStatus - " & "creating Function BAPI_BUS2001_SET_STATUS")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2001_SET_STATUS")
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_GetStatus(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {"PROJECT_DEFINITION"}
        Dim aTables As String() = {"E_SYSTEM_STATUS", "E_USER_STATUS"}
        Try
            log.Debug("getMeta_GetStatus - " & "creating Function BAPI_BUS2001_GET_STATUS")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2001_GET_STATUS")
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function createSingle(pData As TSAP_ProjectData, Optional pOKMsg As String = "OK") As String
        createSingle = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2001_CREATE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oRETURN.Clear()
            oEXTENSIONIN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aProjectinfo.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the extension values - needs dynamic implementation
            Dim aProject As String
            If pData.aExtinfo.aTDataRecCol.Count > 0 Then
                Dim aSAPFormat As New SAPFormat(aIntPar)
                Dim aCustFields As IEnumerable(Of String)
                aProject = pData.getProject()
                aCustFields = fillCustomerFields(aSAPFormat.uneditProj(aProject, 18), pData.aExtinfo)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_PROJECT_DEFINITION")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields.ElementAt(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields.ElementAt(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields.ElementAt(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields.ElementAt(3))
            End If
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                createSingle = createSingle & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    createSingle = createSingle & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            createSingle = If(createSingle = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & createSingle, "Error" & createSingle))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
            createSingle = "Error: Exception in createSingle"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function changeSingle(pData As TSAP_ProjectChgData, Optional pOKMsg As String = "OK") As String
        changeSingle = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2001_CHANGE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oRETURN.Clear()
            oEXTENSIONIN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aProjectinfo.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the extension values - needs dynamic implementation
            Dim aProject As String
            If pData.aExtinfo.aTDataRecCol.Count > 0 Then
                Dim aSAPFormat As New SAPFormat(aIntPar)
                Dim aCustFields As IEnumerable(Of String)
                aProject = pData.getProject()
                aCustFields = fillCustomerFields(aSAPFormat.uneditProj(aProject, 18), pData.aExtinfo)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_PROJECT_DEFINITION")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields.ElementAt(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields.ElementAt(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields.ElementAt(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields.ElementAt(3))
            End If
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                changeSingle = changeSingle & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    changeSingle = changeSingle & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            changeSingle = If(changeSingle = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & changeSingle, "Error" & changeSingle))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
            changeSingle = "Error: Exception in changeSingle"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Function fillCustomerFields(pProject As String, pExtInfo As TDataRec) As IEnumerable(Of String)
        Dim aSAPFormat As New SAPFormat(aIntPar)
        Dim aExtension As New SAPCommon.SapExtension(aIntPar)
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pExtInfo.aTDataRecCol
            aExtension.addField(aTStrRec)
        Next
        aExtension.addString(pProject, 0, 24)
        fillCustomerFields = aExtension.getArray()
    End Function

    Public Function SetStatus(pData As TSAP_ProjectStatusData, Optional pOKMsg As String = "OK") As String
        SetStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2001_SET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oE_RESULT As IRfcTable = oRfcFunction.GetTable("E_RESULT")
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
            SetStatus = "Error: Exception in SetStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function GetStatus(pData As TSAP_ProjectStatusData, Optional pOKMsg As String = "OK") As String
        GetStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2001_GET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oE_SYSTEM_STATUS As IRfcTable = oRfcFunction.GetTable("E_SYSTEM_STATUS")
            Dim oE_USER_STATUS As IRfcTable = oRfcFunction.GetTable("E_USER_STATUS")
            oE_SYSTEM_STATUS.Clear()
            oE_USER_STATUS.Clear()

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
            End If
            GetStatus = If(GetStatus = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & GetStatus, "Error" & GetStatus))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinitionPI")
            GetStatus = "Error: Exception in GetStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
