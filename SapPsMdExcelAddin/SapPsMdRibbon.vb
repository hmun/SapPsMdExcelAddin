Imports Microsoft.Office.Tools.Ribbon

Public Class SapPsMdRibbon
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub SapPsMdRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Sub ButtonProjectCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonProjectCreate.Click
        Dim aSapPsMdRibbonProject As New SapPsMdRibbonProject
        If checkCon() = True Then
            aSapPsMdRibbonProject.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonProjectCreate_Click")
        End If
    End Sub

    Private Sub ButtonProjectChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonProjectChange.Click
        Dim aSapPsMdRibbonProject As New SapPsMdRibbonProject
        If checkCon() = True Then
            aSapPsMdRibbonProject.exec(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonProjectCreate_Click")
        End If
    End Sub

    Private Sub ButtonProjectSetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonProjectSetStatus.Click
        Dim aSapPsMdRibbonProject As New SapPsMdRibbonProject
        If checkCon() = True Then
            aSapPsMdRibbonProject.Status(pSapCon:=aSapCon, pMode:="Set")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonProjectSetStatus_Click")
        End If
    End Sub

    Private Sub ButtonProjectGetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonProjectGetStatus.Click
        Dim aSapPsMdRibbonProject As New SapPsMdRibbonProject
        If checkCon() = True Then
            aSapPsMdRibbonProject.Status(pSapCon:=aSapCon, pMode:="Get")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonProjectSetStatus_Click")
        End If
    End Sub

    Private Sub ButtonWbsCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWbsCreate.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonWbsCreate_Click")
        End If
    End Sub
    Private Sub ButtonWbsCreateSingleMode_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWbsCreateSingleMode.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.exec(pSapCon:=aSapCon, pMode:="CreateSingle")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonWbsCreate_Click")
        End If
    End Sub


    Private Sub ButtonWbsChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWbsChange.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.exec(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonWbsChange_Click")
        End If
    End Sub

    Private Sub ButtonWbsSettlementCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWbsSettlementCreate.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.exec_settle(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonWbsSettlementCreate_Click")
        End If
    End Sub

    Private Sub ButtonWBSGetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWBSGetStatus.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.Status(pSapCon:=aSapCon, pMode:="Get")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonWBSGetStatus_Click")
        End If
    End Sub

    Private Sub ButtonWBSSetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWBSSetStatus.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.Status(pSapCon:=aSapCon, pMode:="Set")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonWBSSetStatus_Click")
        End If
    End Sub

    Private Sub ButtonNetworkCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNetworkCreate.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNetworkCreate_Click")
        End If
    End Sub

    Private Sub ButtonNetworkChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNetworkChange.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.exec(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNetworkChange_Click")
        End If
    End Sub
    Private Sub ButtonNetworkGetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNetworkGetStatus.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.StatusNetwork(pSapCon:=aSapCon, pMode:="Get")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNetworkGetStatus_Click")
        End If
    End Sub

    Private Sub ButtonNetworkSetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNetworkSetStatus.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.StatusNetwork(pSapCon:=aSapCon, pMode:="Set")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNetworkSetStatus_Click")
        End If
    End Sub

    Private Sub ButtonNWACreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNWACreate.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.execNWA(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNWACreate_Click")
        End If
    End Sub

    Private Sub ButtonNWAChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNWAChange.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.execNWA(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNWAChange_Click")
        End If
    End Sub

    Private Sub ButtonNWAGetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNWAGetStatus.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.StatusNWA(pSapCon:=aSapCon, pMode:="Get")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNWAGetStatus_Click")
        End If
    End Sub

    Private Sub ButtonNWASetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNWASetStatus.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.StatusNWA(pSapCon:=aSapCon, pMode:="Set")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNWASetStatus_Click")
        End If
    End Sub

    Private Sub ButtonNWAECreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNWAECreate.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.execNWAE(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNWAECreate_Click")
        End If
    End Sub

    Private Sub ButtonNWAEChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNWAEChange.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.execNWAE(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNWAEChange_Click")
        End If
    End Sub

    Private Sub ButtonCompCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCompCreate.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.execCOMP(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCompCreate_Click")
        End If
    End Sub

    Private Sub ButtonCompChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCompChange.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.execCOMP(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCompChange_Click")
        End If
    End Sub

    Private Sub ButtonCombProjCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCombProjCreate.Click
        Dim aSapPsMdRibbonProject As New SapPsMdRibbonProject
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonProject.exec(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonWbs.exec(pSapCon:=aSapCon, pMode:="Create")
            MsgBox("Complete! Check Messages in the Project and WBS sheets", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap ButtonCombProjCreate_Click")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonProjectCreate_Click")
        End If
    End Sub

    Private Sub ButtonCombNetwCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCombNetwCreate.Click
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonNetwork.exec(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.execNWA(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.execNWAE(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.execCOMP(pSapCon:=aSapCon, pMode:="Create")
            MsgBox("Complete! Check Messages in the Network, NetworkActivity, NetwActElement, Component sheets", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap ButtonCombProjCreate_Click")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNetworkCreate_Click")
        End If

    End Sub

    Private Sub ButtonCombAllCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCombAllCreate.Click
        Dim aSapPsMdRibbonProject As New SapPsMdRibbonProject
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        Dim aSapPsMdRibbonNetwork As New SapPsMdRibbonNetwork
        If checkCon() = True Then
            aSapPsMdRibbonProject.exec(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonWbs.exec(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.exec(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.execNWA(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.execNWAE(pSapCon:=aSapCon, pMode:="Create")
            aSapPsMdRibbonNetwork.execCOMP(pSapCon:=aSapCon, pMode:="Create")
            MsgBox("Complete! Check Messages in all sheets", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap ButtonCombAllCreate_Click")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonProjectCreate_Click")
        End If

    End Sub
End Class
