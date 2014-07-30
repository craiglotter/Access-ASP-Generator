Public Class User_Input_Prompt
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal labeltext As String)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Label1.Text = labeltext
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents input As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents accept_button As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(User_Input_Prompt))
        Me.input = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cancel_button = New System.Windows.Forms.Button
        Me.accept_button = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'input
        '
        Me.input.Location = New System.Drawing.Point(16, 48)
        Me.input.Name = "input"
        Me.input.Size = New System.Drawing.Size(240, 20)
        Me.input.TabIndex = 0
        Me.input.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(408, 32)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Please enter a valid input value."
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(344, 48)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 2
        Me.cancel_button.Text = "Cancel"
        '
        'accept_button
        '
        Me.accept_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.accept_button.Location = New System.Drawing.Point(264, 48)
        Me.accept_button.Name = "accept_button"
        Me.accept_button.TabIndex = 1
        Me.accept_button.Text = "Accept"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(16, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(408, 32)
        Me.Panel1.TabIndex = 4
        '
        'User_Input_Prompt
        '
        Me.AcceptButton = Me.accept_button
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cancel_button
        Me.ClientSize = New System.Drawing.Size(434, 88)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.accept_button)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.input)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "User_Input_Prompt"
        Me.Text = "User Input Required"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
