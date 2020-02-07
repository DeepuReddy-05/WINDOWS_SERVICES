<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
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
        Me.ServiceProcessInstaller_Shopify_Description_Update = New System.ServiceProcess.ServiceProcessInstaller()
        Me.ServiceInstaller_Shopify_Description_Update = New System.ServiceProcess.ServiceInstaller()
        '
        'ServiceProcessInstaller_Shopify_Description_Update
        '
        Me.ServiceProcessInstaller_Shopify_Description_Update.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessInstaller_Shopify_Description_Update.Password = Nothing
        Me.ServiceProcessInstaller_Shopify_Description_Update.Username = Nothing
        '
        'ServiceInstaller_Shopify_Description_Update
        '
        Me.ServiceInstaller_Shopify_Description_Update.DelayedAutoStart = True
        Me.ServiceInstaller_Shopify_Description_Update.ServiceName = "OmegaCube_Shopify_Description_Update"
        Me.ServiceInstaller_Shopify_Description_Update.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstaller_Shopify_Description_Update, Me.ServiceInstaller_Shopify_Description_Update})

    End Sub

    Friend WithEvents ServiceProcessInstaller_Shopify_Description_Update As ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServiceInstaller_Shopify_Description_Update As ServiceProcess.ServiceInstaller
End Class
