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

End Class
