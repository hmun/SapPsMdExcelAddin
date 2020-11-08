Public Class SapPsMdExcelAddin

    Private Sub SapPsMdExcelAddin_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapPsMdExcelAddin_Shutdown() Handles Me.Shutdown

    End Sub

End Class
