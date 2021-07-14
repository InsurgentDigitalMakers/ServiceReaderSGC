Imports System.ComponentModel
Imports System.Configuration.Install

Public Class ProjectInstaller

    Public Sub New()
        MyBase.New()

        'El Diseñador de componentes requiere esta llamada.
        InitializeComponent()

        'Agregue el código de inicialización después de llamar a InitializeComponent

    End Sub

    Private Sub ServiceProcessInstaller1_AfterInstall(sender As Object, e As InstallEventArgs) Handles ServiceProcessInstallerReaderSGC.AfterInstall

    End Sub
End Class
