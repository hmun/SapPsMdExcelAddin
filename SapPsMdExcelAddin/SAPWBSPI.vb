' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPWBSPI
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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        End Try
    End Sub

    Public Function createMultiple(pData As TSAP_WbsData, Optional pOKMsg As String = "OK") As String
        createMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_CREATE_MULTI")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_WBS_ELEMENT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oRETURN.Clear()
            oIT_WBS_ELEMENT.Clear()
            oEXTENSIONIN.Clear()

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
                Dim oIT_WBS_ELEMENTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_WBS_ELEMENT"
                            If Not oIT_WBS_ELEMENTAppended Then
                                oIT_WBS_ELEMENT.Append()
                                oIT_WBS_ELEMENTAppended = True
                            End If
                            oIT_WBS_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' Fill Extension fields
            Dim oEXTENSIONINAppended As Boolean = False
            For Each aKvP In pData.aExt.aTDataDic
                aTDataRec = aKvP.Value
                Dim aCustFields As Object
                aCustFields = fillCustomerFields(aTDataRec)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_WBS_ELEMENT")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields(3))
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                createMultiple = createMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    createMultiple = createMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            createMultiple = If(createMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & createMultiple, "Error" & createMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            createMultiple = "Error: Exception in createMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    ' TBD: This Function needs to be reimplemented !!!
    Function fillCustomerFields(pExtInfo As TDataRec) As Object
        Dim aArray(3) As String
        Dim aRetArray As Object
        Dim aSAPFormat As New SAPFormat(aIntPar)
        ' WBS
        aRetArray = fillEXTENSIONIN(aArray, aSAPFormat.pspid(pExtInfo.getWbs, 18), 0, 1, 24, True)

        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getWbsZZ_REL("1"), 0, 25, 12, False)
        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getWbsZZ_REL("2"), 0, 37, 12, False)
        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getWbsZZ_REL("3"), 0, 49, 12, False)
        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getWbsZZ_REL("4"), 0, 61, 12, False)

        fillCustomerFields = aRetArray
    End Function

    Function fillEXTENSIONIN(pArray As Object, pValue As String, pInd As Integer, pStart As Integer, pLen As Integer, pClear As Boolean) As Object
        Dim aArray(3) As String
        Dim eStr As String
        Dim tmpStr As String
        eStr = "                                                                                                                                                                                                                                                "
        For i = 0 To 3
            If pClear = True Then
                aArray(i) = eStr
            Else
                aArray(i) = pArray(i)
            End If
        Next i
        tmpStr = Left(aArray(pInd), pStart - 1)
        tmpStr = tmpStr & pValue
        tmpStr = tmpStr & Left(eStr, pLen - Len(pValue))
        tmpStr = tmpStr & Right(aArray(pInd), Len(aArray(pInd)) - Len(tmpStr))
        aArray(pInd) = tmpStr
        fillEXTENSIONIN = aArray
    End Function

    Public Function createSettlementRule(pData As TSAP_WbsSettleData, Optional pOKMsg As String = "OK") As String
        createSettlementRule = ""
        Dim aSAPFormat As New SAPFormat(aIntPar)
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZPS_KSRG_WBS")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            ' use local Version of the SapFormat.pspid (the common does not support the mask strings)
            If pData.aHdrRec.aTDataRecCol.Count <> 3 Then
                createSettlementRule = pOKMsg & "; not relevant"
                Exit Function
            End If
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    If String.IsNullOrEmpty(aTStrRec.Value) Then
                        createSettlementRule = pOKMsg & "; not relevant"
                        Exit Function
                    Else
                        If Left(aTStrRec.Format, 1) = "P" Then
                            oRfcFunction.SetValue(aTStrRec.Fieldname, aSAPFormat.pspid(aTStrRec.Value, 18))
                        Else
                            oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                        End If
                    End If

                End If
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                createSettlementRule = createSettlementRule & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            createSettlementRule = If(createSettlementRule = "", pOKMsg, If(aErr = False, pOKMsg & createSettlementRule, "Error" & createSettlementRule))
        Catch SapEx As SAP.Middleware.Connector.RfcAbapMessageException
            createSettlementRule = "Error; " & SapEx.AbapMessageType & "-" & SapEx.AbapMessageClass & "-" & SapEx.AbapMessageNumber & ": " & SapEx.Message
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            createSettlementRule = "Error: Exception in createSettlementRule"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function
End Class
