' Copyright 2020 Hermann Mundprecht
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

End Class
