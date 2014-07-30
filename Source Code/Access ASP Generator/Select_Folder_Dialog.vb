Public Class Select_Folder_Dialog
    Inherits System.Windows.Forms.Form

    Public old_drive As String
    Public old_path As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        old_drive = DriveListBox1.Drive
        old_path = DirListBox1.Path
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
    Friend WithEvents DirListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DirListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.DriveListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'DirListBox1
        '
        Me.DirListBox1.IntegralHeight = False
        Me.DirListBox1.Location = New System.Drawing.Point(56, 48)
        Me.DirListBox1.Name = "DirListBox1"
        Me.DirListBox1.Size = New System.Drawing.Size(280, 96)
        Me.DirListBox1.TabIndex = 0
        '
        'DriveListBox1
        '
        Me.DriveListBox1.Location = New System.Drawing.Point(56, 16)
        Me.DriveListBox1.Name = "DriveListBox1"
        Me.DriveListBox1.Size = New System.Drawing.Size(280, 21)
        Me.DriveListBox1.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(56, 152)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(280, 20)
        Me.TextBox1.TabIndex = 2
        Me.TextBox1.Text = "D:\Work Projects\Manga_CD_List_Generator"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.Location = New System.Drawing.Point(104, 184)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Proceed"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Drive:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 24)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Folder:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 23)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Path:"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Gainsboro
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(200, 184)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Cancel"
        '
        'Select_Folder_Dialog
        '
        Me.AcceptButton = Me.Button1
        Me.AccessibleDescription = "Select_Folder_Dialog"
        Me.AccessibleName = "Select_Folder_Dialog"
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(392, 230)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button2, Me.Label3, Me.Label2, Me.Label1, Me.Button1, Me.TextBox1, Me.DriveListBox1, Me.DirListBox1})
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(400, 257)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(400, 257)
        Me.Name = "Select_Folder_Dialog"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select a Folder"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    Private Sub DirListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox1.SelectedIndexChanged
        TextBox1.Text = DirListBox1.Path
        old_path = DirListBox1.Path
    End Sub

    Private Sub DirListBox1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox1.DoubleClick
        TextBox1.Text = DirListBox1.Path
        old_path = DirListBox1.Path
    End Sub

    Private Sub DriveListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox1.SelectedIndexChanged
        On Error GoTo ErrorHandler
        DirListBox1.Path = DriveListBox1.Drive
        TextBox1.Text = DirListBox1.Path
        old_drive = DriveListBox1.Drive
        old_path = DirListBox1.Path
        Exit Sub
ErrorHandler:
        MsgBox("There is no disc in the Drive.")
        DriveListBox1.Drive = old_drive
        DirListBox1.Path = old_path
        TextBox1.Text = old_path
        Resume Next
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        Me.Hide()
    End Sub

    Private Sub Select_Folder_Dialog_UnLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        TextBox1.Text = ""
    End Sub

End Class
