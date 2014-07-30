Public Class Screensaver
    Inherits System.Windows.Forms.Form
    Dim mousemoveaccept As Boolean = False
    Dim mousemovecount As Integer = 0
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mousemoveaccept = False
        mousemovecount = 0
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Screensaver))
        Me.Button1 = New System.Windows.Forms.Button
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(224, 232)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "QUIT"
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(0, 0)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(292, 266)
        Me.AxShockwaveFlash1.TabIndex = 1
        '
        'Screensaver
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Screensaver"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Screensaver"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal Message As String)
        Try
            MsgBox("Sorry, but the following error has been trapped by Simple Screensaver 1.0: " & vbCrLf & Message, MsgBoxStyle.Exclamation, "Simple Screensaver 1.0 Error")
        Catch ex As Exception
            MsgBox("Sorry, but the following error has been trapped by Simple Screensaver 1.0: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Simple Screensaver 1.0 Error")
        End Try
    End Sub

    Private Sub Exit_Routine()
        Try
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Loader()
        Try
            AxShockwaveFlash1.Movie = Application.StartupPath & "\Simple_Screensaver.swf"
            AxShockwaveFlash1.EmbedMovie = True
            AxShockwaveFlash1.Menu = False
            AxShockwaveFlash1.Enabled = False
            Dim point As New System.Drawing.Point(0, 0)
            Button1.Top = Screen.GetBounds(point).Height - (Button1.Height + 15)
            Button1.Left = Screen.GetBounds(point).Width - (Button1.Width + 10)
            'Button1.Focus()
            Button1.Visible = False
            mousemoveaccept = True
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, MyBase.Click
        Try
            Exit_Routine()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Button1.MouseMove
        Try
            mousemovecount = mousemovecount + 1
            If mousemoveaccept = True And mousemovecount = 5 Then
                Exit_Routine()
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Capture_Keypress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress, Button1.KeyPress
        Try
            If Not IsNothing(e.KeyChar) Then
                e.Handled = True
                Exit_Routine()
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Screensaver_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Loader()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub




End Class
