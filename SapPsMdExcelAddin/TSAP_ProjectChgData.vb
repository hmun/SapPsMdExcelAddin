' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TSAP_ProjectChgData

    Public aProjectinfo As TDataRec
    Public aExtinfo As TDataRec

    Private Project_Fields_Chg() As String = {"PROJECT_DEFINITION", "DESCRIPTION", "MASK_ID", "WBS_STATUS_PROFILE", "RESPONSIBLE_NO", "APPLICANT_NO", "COMPANY_CODE", "BUSINESS_AREA", "PROFIT_CTR", "PROJECT_CURRENCY", "PROJECT_CURRENCY_ISO", "START", "FINISH", "PLANT", "CALENDAR", "PLAN_BASIC", "PLAN_FCST", "TIME_UNIT", "TIME_UNIT_ISO", "NETWORK_PROFILE", "BUDGET_PROFILE", "PROJECT_STOCK", "OBJECTCLASS", "STATISTICAL", "TAXJURCODE", "INTEREST_PROF", "WBS_SCHED_PROFILE", "INVEST_PROFILE", "RES_ANAL_KEY", "PLAN_PROFILE", "PLANINTEGRATED", "VALUATION_SPEC_STOCK", "SIMULATION_PROFILE", "GROUPING_INDICATOR", "LOCATION", "PARTNER_PROFILE", "VENTURE", "REC_IND", "EQUITY_TYP", "JV_OTYPE", "JV_JIBCL", "JV_JIBSA", "SCHED_SCENARIO", "FCST_START", "FCST_FINISH", "FUNC_AREA", "SALESORG", "DISTR_CHAN", "DIVISION", "DLI_PROFILE"}
    Private Project_Fields_Upd() As String = {"DESCRIPTION", "MASK_ID", "WBS_STATUS_PROFILE", "RESPONSIBLE_NO", "APPLICANT_NO", "COMPANY_CODE", "BUSINESS_AREA", "PROFIT_CTR", "PROJECT_CURRENCY", "PROJECT_CURRENCY_ISO", "START", "FINISH", "PLANT", "CALENDAR", "PLAN_BASIC", "PLAN_FCST", "TIME_UNIT", "TIME_UNIT_ISO", "NETWORK_PROFILE", "BUDGET_PROFILE", "PROJECT_STOCK", "OBJECTCLASS", "STATISTICAL", "TAXJURCODE", "INTEREST_PROF", "WBS_SCHED_PROFILE", "INVEST_PROFILE", "RES_ANAL_KEY", "PLAN_PROFILE", "PLANINTEGRATED", "VALUATION_SPEC_STOCK", "SIMULATION_PROFILE", "GROUPING_INDICATOR", "LOCATION", "PARTNER_PROFILE", "VENTURE", "REC_IND", "EQUITY_TYP", "JV_OTYPE", "JV_JIBCL", "JV_JIBSA", "SCHED_SCENARIO", "FCST_START", "FCST_FINISH", "FUNC_AREA", "SALESORG", "DISTR_CHAN", "DIVISION", "DLI_PROFILE"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sProj As String = "I_PROJECT_DEFINITION"
    Private Const sProj_Upd As String = "I_PROJECT_DEFINITION_UPD"

    Private aUseAsEmpty As String = "#"

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
        aUseAsEmpty = If(aIntPar.value("GEN", "USEASEMPTY") <> "", aIntPar.value("GEN", "USEASEMPTY"), "#")
    End Sub

    Public Function fillProjectinfo(pData As TData) As Boolean
        aProjectinfo = New TDataRec(aIntPar)
        aExtinfo = New TDataRec(aIntPar)
        Dim aFirstRec As New TDataRec(aIntPar)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewProjectinfo As New TDataRec(aIntPar)
        Dim aNewExtinfo As New TDataRec(aIntPar)
        aFirstRec = pData.getFirstRecord()
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        If Not IsNothing(aFirstRec) Then
            For Each aTStrRec In aFirstRec.aTDataRecCol
                If valid_Proj_Field(aTStrRec) Then
                    aNewProjectinfo.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pUseAsEmpty:=aUseAsEmpty)
                    If valid_Proj_Field_Upd(aTStrRec) Then
                        aNewProjectinfo.setValues(sProj_Upd & "-" & aTStrRec.Fieldname, "X", "", "", pUseAsEmpty:=aUseAsEmpty)
                    End If
                ElseIf valid_Ext_Field(aTStrRec) Then
                    aNewExtinfo.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pUseAsEmpty:=aUseAsEmpty)
                End If
            Next
        End If
        aProjectinfo = aNewProjectinfo
        aExtinfo = aNewExtinfo
        fillProjectinfo = True
    End Function

    Public Function valid_Proj_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Proj_Field = False
        If pTStrRec.Strucname = "I_PROJECT_DEFINITION" Then
            valid_Proj_Field = isInArray(pTStrRec.Fieldname, Project_Fields_Chg)
        End If
    End Function

    Public Function valid_Proj_Field_Upd(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Proj_Field_Upd = False
        If pTStrRec.Strucname = "I_PROJECT_DEFINITION" Then
            valid_Proj_Field_Upd = isInArray(pTStrRec.Fieldname, Project_Fields_Upd)
        End If
    End Function

    Public Function valid_Ext_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aValExtString As String = If(aIntPar.value("PROJ_STR", "VALEXT") <> "", aIntPar.value("PROJ_STR", "VALEXT"), "")
        valid_Ext_Field = False
        If pTStrRec.Strucname = aValExtString Then
            valid_Ext_Field = True
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Function getProject() As String
        Dim aTStrRec As SAPCommon.TStrRec
        getProject = ""
        For Each aTStrRec In aProjectinfo.aTDataRecCol
            If aTStrRec.Fieldname = "PROJECT_DEFINITION" Then
                getProject = aTStrRec.Value
                Exit Function
            End If
        Next
    End Function

    Public Sub dumpProjectinfo()
        Dim dumpHd As String = If(aIntPar.value("PROJ_DBG", "DUMPDATA") <> "", aIntPar.value("PROJ_DBG", "DUMPDATA"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapPsMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpProjectinfo - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the PROJ_DBG-DUMPDATA Parameter",
                MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
                Exit Sub
            End Try
            log.Debug("dumpProjectinfo - " & "dumping to " & dumpHd)
            ' clear the Projectinfo
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Projectinfo
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aProjectinfo.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
            ' dump the aExtinfo
            aFieldArray = {}
            aValueArray = {}
            For Each aTStrRec In aExtinfo.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(3, 1), aDWS.Cells(3, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(4, 1), aDWS.Cells(4, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

End Class
