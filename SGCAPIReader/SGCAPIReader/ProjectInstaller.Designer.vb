<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    'Se puede modificar usando el Diseñador de componentes.
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ServiceProcessInstallerReaderSGC = New System.ServiceProcess.ServiceProcessInstaller()
        Me.ServiceInstaller1 = New System.ServiceProcess.ServiceInstaller()
        '
        'ServiceProcessInstallerReaderSGC
        '
        Me.ServiceProcessInstallerReaderSGC.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessInstallerReaderSGC.Password = Nothing
        Me.ServiceProcessInstallerReaderSGC.Username = Nothing
        '
        'ServiceInstaller1
        '
        Me.ServiceInstaller1.Description = "Servicio de Lectura de WEB API para cargar datos externos al Sistema SGC"
        Me.ServiceInstaller1.DisplayName = "ServiceSGCReader"
        Me.ServiceInstaller1.ServiceName = "ServiceSGCReader"
        Me.ServiceInstaller1.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstallerReaderSGC, Me.ServiceInstaller1})

    End Sub

    Friend WithEvents ServiceProcessInstallerReaderSGC As ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServiceInstaller1 As ServiceProcess.ServiceInstaller
End Class
