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
            Dim aCustFields As Object
            If pData.aExtinfo.aTDataRecCol.Count > 0 Then
                Dim aSAPFormat As New SAPFormat(aIntPar)
                aProject = pData.getProject()
                aCustFields = fillCustomerFields(aSAPFormat.uneditProj(aProject, 18), pData.aExtinfo)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_PROJECT_DEFINITION")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields(3))
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

    ' TBD: This Function needs to be reimplemented !!!
    Function fillCustomerFields(pProject As String, pExtInfo As TDataRec) As Object
        Dim aArray(3) As String
        Dim aRetArray As Object
        Dim aSAPFormat As New SAPFormat(aIntPar)
        ' Project
        aRetArray = fillEXTENSIONIN(aArray, pProject, 0, 1, 24, True)

        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getProjZZ_REL("1"), 0, 25, 12, False)
        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getProjZZ_REL("2"), 0, 37, 12, False)
        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getProjZZ_REL("3"), 0, 49, 12, False)
        aRetArray = fillEXTENSIONIN(aRetArray, pExtInfo.getProjZZ_REL("4"), 0, 61, 12, False)

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

End Class
