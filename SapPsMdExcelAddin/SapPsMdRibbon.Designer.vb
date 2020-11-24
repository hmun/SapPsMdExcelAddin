Partial Class SapPsMdRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapPsMdRibbon))
        Me.SapPsMd = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ButtonCombAllCreate = Me.Factory.CreateRibbonButton
        Me.ButtonCombProjCreate = Me.Factory.CreateRibbonButton
        Me.ButtonCombNetwCreate = Me.Factory.CreateRibbonButton
        Me.SAPProject = Me.Factory.CreateRibbonGroup
        Me.ButtonProjectCreate = Me.Factory.CreateRibbonButton
        Me.ButtonWbsCreate = Me.Factory.CreateRibbonButton
        Me.SAPNetwork = Me.Factory.CreateRibbonGroup
        Me.ButtonNetworkCreate = Me.Factory.CreateRibbonButton
        Me.ButtonNWACreate = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.ButtonNWAECreate = Me.Factory.CreateRibbonButton
        Me.ButtonCompCreate = Me.Factory.CreateRibbonButton
        Me.SapPsMdLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.ButtonWbsSettlementCreate = Me.Factory.CreateRibbonButton
        Me.SapPsMd.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SAPProject.SuspendLayout()
        Me.SAPNetwork.SuspendLayout()
        Me.SapPsMdLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapPsMd
        '
        Me.SapPsMd.Groups.Add(Me.Group1)
        Me.SapPsMd.Groups.Add(Me.SAPProject)
        Me.SapPsMd.Groups.Add(Me.SAPNetwork)
        Me.SapPsMd.Groups.Add(Me.SapPsMdLogon)
        Me.SapPsMd.Label = "SAP PS Md"
        Me.SapPsMd.Name = "SapPsMd"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonCombAllCreate)
        Me.Group1.Items.Add(Me.ButtonCombProjCreate)
        Me.Group1.Items.Add(Me.ButtonCombNetwCreate)
        Me.Group1.Label = "PS Combined"
        Me.Group1.Name = "Group1"
        '
        'ButtonCombAllCreate
        '
        Me.ButtonCombAllCreate.Image = CType(resources.GetObject("ButtonCombAllCreate.Image"), System.Drawing.Image)
        Me.ButtonCombAllCreate.Label = "Create All"
        Me.ButtonCombAllCreate.Name = "ButtonCombAllCreate"
        Me.ButtonCombAllCreate.ShowImage = True
        '
        'ButtonCombProjCreate
        '
        Me.ButtonCombProjCreate.Image = CType(resources.GetObject("ButtonCombProjCreate.Image"), System.Drawing.Image)
        Me.ButtonCombProjCreate.Label = "Create Project All"
        Me.ButtonCombProjCreate.Name = "ButtonCombProjCreate"
        Me.ButtonCombProjCreate.ShowImage = True
        '
        'ButtonCombNetwCreate
        '
        Me.ButtonCombNetwCreate.Image = CType(resources.GetObject("ButtonCombNetwCreate.Image"), System.Drawing.Image)
        Me.ButtonCombNetwCreate.Label = "Create Network All"
        Me.ButtonCombNetwCreate.Name = "ButtonCombNetwCreate"
        Me.ButtonCombNetwCreate.ShowImage = True
        '
        'SAPProject
        '
        Me.SAPProject.Items.Add(Me.ButtonProjectCreate)
        Me.SAPProject.Items.Add(Me.ButtonWbsCreate)
        Me.SAPProject.Items.Add(Me.ButtonWbsSettlementCreate)
        Me.SAPProject.Label = "PS Project"
        Me.SAPProject.Name = "SAPProject"
        '
        'ButtonProjectCreate
        '
        Me.ButtonProjectCreate.Image = CType(resources.GetObject("ButtonProjectCreate.Image"), System.Drawing.Image)
        Me.ButtonProjectCreate.Label = "Create Project"
        Me.ButtonProjectCreate.Name = "ButtonProjectCreate"
        Me.ButtonProjectCreate.ShowImage = True
        '
        'ButtonWbsCreate
        '
        Me.ButtonWbsCreate.Image = CType(resources.GetObject("ButtonWbsCreate.Image"), System.Drawing.Image)
        Me.ButtonWbsCreate.Label = "Create WBS"
        Me.ButtonWbsCreate.Name = "ButtonWbsCreate"
        Me.ButtonWbsCreate.ShowImage = True
        '
        'SAPNetwork
        '
        Me.SAPNetwork.Items.Add(Me.ButtonNetworkCreate)
        Me.SAPNetwork.Items.Add(Me.ButtonNWACreate)
        Me.SAPNetwork.Items.Add(Me.Separator1)
        Me.SAPNetwork.Items.Add(Me.ButtonNWAECreate)
        Me.SAPNetwork.Items.Add(Me.ButtonCompCreate)
        Me.SAPNetwork.Label = "PS Network"
        Me.SAPNetwork.Name = "SAPNetwork"
        '
        'ButtonNetworkCreate
        '
        Me.ButtonNetworkCreate.Image = CType(resources.GetObject("ButtonNetworkCreate.Image"), System.Drawing.Image)
        Me.ButtonNetworkCreate.Label = "Create Network"
        Me.ButtonNetworkCreate.Name = "ButtonNetworkCreate"
        Me.ButtonNetworkCreate.ShowImage = True
        '
        'ButtonNWACreate
        '
        Me.ButtonNWACreate.Image = CType(resources.GetObject("ButtonNWACreate.Image"), System.Drawing.Image)
        Me.ButtonNWACreate.Label = "Create NetworkActivities"
        Me.ButtonNWACreate.Name = "ButtonNWACreate"
        Me.ButtonNWACreate.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'ButtonNWAECreate
        '
        Me.ButtonNWAECreate.Image = CType(resources.GetObject("ButtonNWAECreate.Image"), System.Drawing.Image)
        Me.ButtonNWAECreate.Label = "Create NetwActElements"
        Me.ButtonNWAECreate.Name = "ButtonNWAECreate"
        Me.ButtonNWAECreate.ShowImage = True
        '
        'ButtonCompCreate
        '
        Me.ButtonCompCreate.Image = CType(resources.GetObject("ButtonCompCreate.Image"), System.Drawing.Image)
        Me.ButtonCompCreate.Label = "Create Components"
        Me.ButtonCompCreate.Name = "ButtonCompCreate"
        Me.ButtonCompCreate.ShowImage = True
        '
        'SapPsMdLogon
        '
        Me.SapPsMdLogon.Items.Add(Me.ButtonLogon)
        Me.SapPsMdLogon.Items.Add(Me.ButtonLogoff)
        Me.SapPsMdLogon.Label = "Logon"
        Me.SapPsMdLogon.Name = "SapPsMdLogon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'ButtonWbsSettlementCreate
        '
        Me.ButtonWbsSettlementCreate.Image = CType(resources.GetObject("ButtonWbsSettlementCreate.Image"), System.Drawing.Image)
        Me.ButtonWbsSettlementCreate.Label = "Create WBS Settlement"
        Me.ButtonWbsSettlementCreate.Name = "ButtonWbsSettlementCreate"
        Me.ButtonWbsSettlementCreate.ShowImage = True
        '
        'SapPsMdRibbon
        '
        Me.Name = "SapPsMdRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapPsMd)
        Me.SapPsMd.ResumeLayout(False)
        Me.SapPsMd.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.SAPProject.ResumeLayout(False)
        Me.SAPProject.PerformLayout()
        Me.SAPNetwork.ResumeLayout(False)
        Me.SAPNetwork.PerformLayout()
        Me.SapPsMdLogon.ResumeLayout(False)
        Me.SapPsMdLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapPsMd As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SAPProject As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonProjectCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SapPsMdLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonWbsCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPNetwork As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonNetworkCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonNWACreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCombAllCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCombProjCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCombNetwCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ButtonNWAECreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCompCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonWbsSettlementCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapPsMdRibbon() As SapPsMdRibbon
        Get
            Return Me.GetRibbon(Of SapPsMdRibbon)()
        End Get
    End Property
End Class
