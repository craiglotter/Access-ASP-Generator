Public Class Change_FileString_Dialog
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents oldstring As System.Windows.Forms.TextBox
    Friend WithEvents newstring As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.oldstring = New System.Windows.Forms.TextBox()
        Me.newstring = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'oldstring
        '
        Me.oldstring.BackColor = System.Drawing.Color.WhiteSmoke
        Me.oldstring.Location = New System.Drawing.Point(16, 40)
        Me.oldstring.Name = "oldstring"
        Me.oldstring.ReadOnly = True
        Me.oldstring.Size = New System.Drawing.Size(264, 20)
        Me.oldstring.TabIndex = 0
        Me.oldstring.Text = ""
        '
        'newstring
        '
        Me.newstring.Location = New System.Drawing.Point(16, 88)
        Me.newstring.Name = "newstring"
        Me.newstring.Size = New System.Drawing.Size(264, 20)
        Me.newstring.TabIndex = 1
        Me.newstring.Text = ""
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.Location = New System.Drawing.Point(112, 128)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Proceed"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Current String:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(216, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "New String:"
        '
        'Change_FileString_Dialog
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(292, 174)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.Label1, Me.Button1, Me.newstring, Me.oldstring})
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(300, 201)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(300, 201)
        Me.Name = "Change_FileString_Dialog"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Change File String"
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    Private Sub Change_FileString_Dialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oldstring.Enabled = False
        newstring.Focus()
    End Sub
End Class
