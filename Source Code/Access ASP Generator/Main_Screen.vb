Imports System
Imports System.IO
Imports System.Collections

Public Class Main_Program

    Inherits System.Windows.Forms.Form

    'Structure declaration
    Public Structure DatabaseColumn
        Public column_name As String
        Public data_type As OleDb.OleDbType
        Public input_type As String
        Public display_column As Boolean
    End Structure

    'Variable declaration

    Private insertdatabase As String
    Private insertfieldkey As String
    Private inserttable As String
    Private insertobjectdescription As String
    Private insertpagetitle As String
    Private insertfilename As String
    Private insertfilename_fullpath As String
    Private template As String

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
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents linsertdatabase As System.Windows.Forms.Label
    Friend WithEvents linserttable As System.Windows.Forms.Label
    Friend WithEvents linsertfieldkey As System.Windows.Forms.Label
    Friend WithEvents linsertobjectdescription As System.Windows.Forms.Label
    Friend WithEvents linsertfilename As System.Windows.Forms.Label
    Friend WithEvents SelectTemplate As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveResult As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ltemplate As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents step01 As System.Windows.Forms.Button
    Friend WithEvents step02 As System.Windows.Forms.Button
    Friend WithEvents step03 As System.Windows.Forms.Button
    Friend WithEvents step05 As System.Windows.Forms.Button
    Friend WithEvents step06 As System.Windows.Forms.Button
    Friend WithEvents step07 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents step04 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents linsertpagetitle As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SelectDatabase As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents LoadJobFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents step09 As System.Windows.Forms.Button

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Program))
        Me.SelectDatabase = New System.Windows.Forms.OpenFileDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.step01 = New System.Windows.Forms.Button
        Me.step02 = New System.Windows.Forms.Button
        Me.step03 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.step04 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.step05 = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.step09 = New System.Windows.Forms.Button
        Me.step06 = New System.Windows.Forms.Button
        Me.step07 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.linsertdatabase = New System.Windows.Forms.Label
        Me.linserttable = New System.Windows.Forms.Label
        Me.linsertobjectdescription = New System.Windows.Forms.Label
        Me.linsertfilename = New System.Windows.Forms.Label
        Me.linsertfieldkey = New System.Windows.Forms.Label
        Me.SelectTemplate = New System.Windows.Forms.OpenFileDialog
        Me.SaveResult = New System.Windows.Forms.SaveFileDialog
        Me.ltemplate = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.linsertpagetitle = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.LoadJobFile = New System.Windows.Forms.OpenFileDialog
        Me.SuspendLayout()
        '
        'SelectDatabase
        '
        Me.SelectDatabase.DefaultExt = "mdb"
        Me.SelectDatabase.Filter = "Microsoft Access Files|*.mdb"
        Me.SelectDatabase.Title = "Select Microsoft Access Database"
        '
        'step01
        '
        Me.step01.Enabled = False
        Me.step01.Location = New System.Drawing.Point(16, 16)
        Me.step01.Name = "step01"
        Me.step01.TabIndex = 1
        Me.step01.Text = "Step 1"
        Me.ToolTip1.SetToolTip(Me.step01, "Select the template file that will be used to generate the ASP page from.")
        '
        'step02
        '
        Me.step02.Enabled = False
        Me.step02.Location = New System.Drawing.Point(16, 40)
        Me.step02.Name = "step02"
        Me.step02.TabIndex = 2
        Me.step02.Text = "Step 2"
        Me.ToolTip1.SetToolTip(Me.step02, "Select the filename for the generated ASP page.")
        '
        'step03
        '
        Me.step03.Enabled = False
        Me.step03.Location = New System.Drawing.Point(16, 64)
        Me.step03.Name = "step03"
        Me.step03.TabIndex = 3
        Me.step03.Text = "Step 3"
        Me.ToolTip1.SetToolTip(Me.step03, "Set the text for the <Title> tag of the generated ASP page.")
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(104, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 24)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Template:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label1, "Template to be used in generating the ASP page.")
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(104, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 24)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "ASP Filename:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label2, "The filename for the generated ASP page.")
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(104, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 24)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Page Title:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label3, "The <Title> value for the generated ASP page.")
        '
        'step04
        '
        Me.step04.Enabled = False
        Me.step04.Location = New System.Drawing.Point(16, 88)
        Me.step04.Name = "step04"
        Me.step04.TabIndex = 4
        Me.step04.Text = "Step 4"
        Me.ToolTip1.SetToolTip(Me.step04, "Set the object description to be used on the page. ")
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(104, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 24)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "Object Description:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label4, "The object description used on the page. For example, if the database table deals" & _
        " with user information, then a good object description would be ""User Account""")
        '
        'step05
        '
        Me.step05.Enabled = False
        Me.step05.Location = New System.Drawing.Point(16, 112)
        Me.step05.Name = "step05"
        Me.step05.TabIndex = 5
        Me.step05.Text = "Step 5"
        Me.ToolTip1.SetToolTip(Me.step05, "Select the Access database for which the ASP page is going to be generated.")
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(104, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 24)
        Me.Label5.TabIndex = 32
        Me.Label5.Text = "Database:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label5, "The Access database for which the ASP page is to be generated.")
        '
        'step09
        '
        Me.step09.Enabled = False
        Me.step09.Location = New System.Drawing.Point(16, 200)
        Me.step09.Name = "step09"
        Me.step09.TabIndex = 10
        Me.step09.Text = "Generate"
        Me.ToolTip1.SetToolTip(Me.step09, "Generate the ASP page.")
        '
        'step06
        '
        Me.step06.Enabled = False
        Me.step06.Location = New System.Drawing.Point(16, 136)
        Me.step06.Name = "step06"
        Me.step06.TabIndex = 6
        Me.step06.Text = "Step 6"
        Me.ToolTip1.SetToolTip(Me.step06, "Select the database table that is going to be affected by the ASP page.")
        '
        'step07
        '
        Me.step07.Enabled = False
        Me.step07.Location = New System.Drawing.Point(16, 160)
        Me.step07.Name = "step07"
        Me.step07.TabIndex = 7
        Me.step07.Text = "Step 7"
        Me.ToolTip1.SetToolTip(Me.step07, "Select the Primary Key on which database statements are going to be based.")
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(104, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 24)
        Me.Label6.TabIndex = 33
        Me.Label6.Text = "Table:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label6, "The database table affected by the ASP page.")
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(104, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 24)
        Me.Label7.TabIndex = 34
        Me.Label7.Text = "Primary Key:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label7, "The primary key on which database statements are going to be based.")
        '
        'linsertdatabase
        '
        Me.linsertdatabase.BackColor = System.Drawing.Color.Transparent
        Me.linsertdatabase.Location = New System.Drawing.Point(208, 112)
        Me.linsertdatabase.Name = "linsertdatabase"
        Me.linsertdatabase.Size = New System.Drawing.Size(520, 24)
        Me.linsertdatabase.TabIndex = 16
        Me.linsertdatabase.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'linserttable
        '
        Me.linserttable.Location = New System.Drawing.Point(208, 136)
        Me.linserttable.Name = "linserttable"
        Me.linserttable.Size = New System.Drawing.Size(520, 24)
        Me.linserttable.TabIndex = 18
        Me.linserttable.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'linsertobjectdescription
        '
        Me.linsertobjectdescription.BackColor = System.Drawing.Color.Transparent
        Me.linsertobjectdescription.Location = New System.Drawing.Point(208, 88)
        Me.linsertobjectdescription.Name = "linsertobjectdescription"
        Me.linsertobjectdescription.Size = New System.Drawing.Size(520, 24)
        Me.linsertobjectdescription.TabIndex = 19
        Me.linsertobjectdescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'linsertfilename
        '
        Me.linsertfilename.BackColor = System.Drawing.Color.Transparent
        Me.linsertfilename.Location = New System.Drawing.Point(208, 40)
        Me.linsertfilename.Name = "linsertfilename"
        Me.linsertfilename.Size = New System.Drawing.Size(520, 24)
        Me.linsertfilename.TabIndex = 20
        Me.linsertfilename.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'linsertfieldkey
        '
        Me.linsertfieldkey.BackColor = System.Drawing.Color.Transparent
        Me.linsertfieldkey.Location = New System.Drawing.Point(208, 160)
        Me.linsertfieldkey.Name = "linsertfieldkey"
        Me.linsertfieldkey.Size = New System.Drawing.Size(520, 24)
        Me.linsertfieldkey.TabIndex = 22
        Me.linsertfieldkey.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SelectTemplate
        '
        Me.SelectTemplate.DefaultExt = "txt"
        Me.SelectTemplate.Filter = "Template Text Files|*.txt"
        Me.SelectTemplate.Title = "Select Template Text File"
        '
        'SaveResult
        '
        Me.SaveResult.DefaultExt = "asp"
        Me.SaveResult.FileName = "Generated1"
        Me.SaveResult.Filter = "ASP files|*.asp"
        Me.SaveResult.Title = "Save Generated ASP File As"
        '
        'ltemplate
        '
        Me.ltemplate.BackColor = System.Drawing.Color.Transparent
        Me.ltemplate.Location = New System.Drawing.Point(208, 16)
        Me.ltemplate.Name = "ltemplate"
        Me.ltemplate.Size = New System.Drawing.Size(520, 24)
        Me.ltemplate.TabIndex = 24
        Me.ltemplate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem11, Me.MenuItem1, Me.MenuItem4, Me.MenuItem9})
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 0
        Me.MenuItem11.Text = "Exit"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem7, Me.MenuItem8, Me.MenuItem6})
        Me.MenuItem1.Text = "Generate"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Simple Create Page"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Simple Remove Page"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Text = "Simple Edit Page"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 3
        Me.MenuItem8.Text = "Simple Display Page"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 4
        Me.MenuItem6.Text = "Custom Page"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5})
        Me.MenuItem4.Text = "Tools"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "Recreate Templates"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 3
        Me.MenuItem9.Text = "About"
        '
        'linsertpagetitle
        '
        Me.linsertpagetitle.BackColor = System.Drawing.Color.Transparent
        Me.linsertpagetitle.Location = New System.Drawing.Point(208, 64)
        Me.linsertpagetitle.Name = "linsertpagetitle"
        Me.linsertpagetitle.Size = New System.Drawing.Size(520, 24)
        Me.linsertpagetitle.TabIndex = 28
        Me.linsertpagetitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(504, 152)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(240, 128)
        Me.PictureBox1.TabIndex = 35
        Me.PictureBox1.TabStop = False
        '
        'Main_Program
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(736, 241)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.linsertpagetitle)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.step07)
        Me.Controls.Add(Me.step06)
        Me.Controls.Add(Me.step05)
        Me.Controls.Add(Me.step04)
        Me.Controls.Add(Me.step03)
        Me.Controls.Add(Me.step02)
        Me.Controls.Add(Me.step01)
        Me.Controls.Add(Me.ltemplate)
        Me.Controls.Add(Me.linsertfieldkey)
        Me.Controls.Add(Me.linsertfilename)
        Me.Controls.Add(Me.linsertobjectdescription)
        Me.Controls.Add(Me.linserttable)
        Me.Controls.Add(Me.linsertdatabase)
        Me.Controls.Add(Me.step09)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.Name = "Main_Program"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Access ASP Generator 1.0"
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Handles clicking on Exit menu option. 
    'Closes the program
    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        'close the current form
        Me.Close()
    End Sub

    'Handles clicking on Tools >> Recreate Templates menu option.
    'Calls the procedure responsible for generating the templates
    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Generate_Templates()
    End Sub

    'Generates predefined hard-coded templates for creating, removing, editing and displaying
    Private Sub Generate_Templates()
        'create instance of Template_Generator class. This class is responsible for the physical template file creation.
        Dim template_gen1 As Template_Generator = New Template_Generator

        'strings that are responsible for reporting success or failure of the template generator procedure
        Dim successreportstring As String
        Dim failedreportstring As String
        successreportstring = ""
        failedreportstring = ""

        'generate simple create page
        If template_gen1.Generate_Simple_Create_Page() = True Then
            successreportstring = successreportstring & vbCrLf & "   " & "Template_Simple_Create_Page.txt"
        Else
            failedreportstring = failedreportstring & vbCrLf & "   " & "Template_Simple_Create_Page.txt"
        End If

        'generate simple remove page
        If template_gen1.Generate_Simple_Remove_Page() = True Then
            successreportstring = successreportstring & vbCrLf & "   " & "Template_Simple_Remove_Page.txt"
        Else
            failedreportstring = failedreportstring & vbCrLf & "   " & "Template_Simple_Remove_Page.txt"
        End If

        'generate simple edit page
        If template_gen1.Generate_Simple_Edit_Page() = True Then
            successreportstring = successreportstring & vbCrLf & "   " & "Template_Simple_Edit_Page.txt"
        Else
            failedreportstring = failedreportstring & vbCrLf & "   " & "Template_Simple_Edit_Page.txt"
        End If

        'generate simple display page
        If template_gen1.Generate_Simple_Display_Page() = True Then
            successreportstring = successreportstring & vbCrLf & "   " & "Template_Simple_Display_Page.txt"
        Else
            failedreportstring = failedreportstring & vbCrLf & "   " & "Template_Simple_Display_Page.txt"
        End If

        'display result message to user
        MsgBox("Success:" & vbCrLf & successreportstring & vbCrLf & vbCrLf & "Failure:" & vbCrLf & failedreportstring, MsgBoxStyle.OKOnly, "Template Generation Results")
    End Sub

    'Handles clicking on Generate >> Custom Page menu option
    'Calls the procedure responsible for generating a custom ASP page
    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        'disables all the step buttons
        Disable_Step_Buttons()
        'clears the variables
        Clear_Variables()
        'calls the procedure responsible for generating a custom ASP page
        Step01_Code("Custom")
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Disable_Step_Buttons()
        Clear_Variables()
        Step01_Code("Simple_Create")
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Disable_Step_Buttons()
        Clear_Variables()
        Step01_Code("Simple_Remove")
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Disable_Step_Buttons()
        Clear_Variables()
        Step01_Code("Simple_Edit")
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        Disable_Step_Buttons()
        Clear_Variables()
        Step01_Code("Simple_Display")
    End Sub

    Private Sub Disable_Step_Buttons()
        step01.Enabled = False
        step02.Enabled = False
        step03.Enabled = False
        step04.Enabled = False
        step05.Enabled = False
        step06.Enabled = False
        step07.Enabled = False
        'step08.Enabled = False
        step09.Enabled = False
    End Sub

    Private Sub Clear_Variables()
        insertdatabase = Nothing
        insertfieldkey = Nothing
        inserttable = Nothing
        insertobjectdescription = Nothing
        insertpagetitle = Nothing
        insertfilename = Nothing
        insertfilename_fullpath = Nothing
        template = Nothing
        linsertdatabase.Text = Nothing
        linsertfieldkey.Text = Nothing
        linserttable.Text = Nothing
        linsertobjectdescription.Text = Nothing
        linsertpagetitle.Text = Nothing
        linsertfilename.Text = Nothing
        ltemplate.Text = Nothing
    End Sub

    Private Sub Step01_Code(ByVal TypeGeneration As String)
        Select Case TypeGeneration
            Case "Custom"
                step01.Enabled = True
                DialogResult = SelectTemplate.ShowDialog()
                If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then
                    template = SelectTemplate.FileName
                    ltemplate.Text = template
                End If
            Case "Simple_Create"
                Dim f, g As FileInfo
                f = New FileInfo(Application.ExecutablePath())
                g = New FileInfo(f.DirectoryName() & "\ASP_Templates\Template_Simple_Create_Page.txt")
                If g.Exists Then
                    template = g.FullName
                    ltemplate.Text = template
                Else
                    MsgBox("Unable to locate the template file ""Template_Simple_Create_Page.txt"". It is suggested you regenerate the template using the ""Recreate Templates"" function found on the Tools menu.", MsgBoxStyle.Exclamation, "Error Locating Template File")
                End If
            Case "Simple_Remove"
                Dim f, g As FileInfo
                f = New FileInfo(Application.ExecutablePath())
                g = New FileInfo(f.DirectoryName() & "\ASP_Templates\Template_Simple_Remove_Page.txt")
                If g.Exists Then
                    template = g.FullName
                    ltemplate.Text = template
                Else
                    MsgBox("Unable to locate the template file ""Template_Simple_Remove_Page.txt"". It is suggested you regenerate the template using the ""Recreate Templates"" function found on the Tools menu.", MsgBoxStyle.Exclamation, "Error Locating Template File")
                End If
            Case "Simple_Edit"
                Dim f, g As FileInfo
                f = New FileInfo(Application.ExecutablePath())
                g = New FileInfo(f.DirectoryName() & "\ASP_Templates\Template_Simple_Edit_Page.txt")
                If g.Exists Then
                    template = g.FullName
                    ltemplate.Text = template
                Else
                    MsgBox("Unable to locate the template file ""Template_Simple_Edit_Page.txt"". It is suggested you regenerate the template using the ""Recreate Templates"" function found on the Tools menu.", MsgBoxStyle.Exclamation, "Error Locating Template File")
                End If

            Case "Simple_Display"
                Dim f, g As FileInfo
                f = New FileInfo(Application.ExecutablePath())
                g = New FileInfo(f.DirectoryName() & "\ASP_Templates\Template_Simple_Display_Page.txt")
                If g.Exists Then
                    template = g.FullName
                    ltemplate.Text = template
                Else
                    MsgBox("Unable to locate the template file ""Template_Simple_Display_Page.txt"". It is suggested you regenerate the template using the ""Recreate Templates"" function found on the Tools menu.", MsgBoxStyle.Exclamation, "Error Locating Template File")
                End If
        End Select
        ToolTip1.SetToolTip(ltemplate, template)
        If Not template = Nothing Then
            step02.Enabled = True
            step02.Focus()
        End If
    End Sub

    Private Sub step01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step01.Click
        Step01_Code("Custom")
    End Sub

    Private Sub Step02_Code()
        DialogResult = SaveResult.ShowDialog()
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then
            insertfilename = SaveResult.FileName
            insertfilename = insertfilename.Remove(0, insertfilename.LastIndexOf("\") + 1)
            insertfilename_fullpath = SaveResult.FileName
            linsertfilename.Text = insertfilename_fullpath
        End If
        ToolTip1.SetToolTip(linsertfilename, insertfilename_fullpath)
        If Not insertfilename_fullpath = Nothing Then
            step03.Enabled = True
            step03.Focus()
        End If
    End Sub

    Private Sub step02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step02.Click
        Step02_Code()
    End Sub

    Private Sub Step03_Code()
        Dim get_input As User_Input_Prompt = New User_Input_Prompt("Please enter the <Title> text to be used for the generated ASP page.")
        If get_input.ShowDialog() = DialogResult.OK Then
            If Not get_input.input.Text = Nothing Then
                insertpagetitle = get_input.input.Text
                linsertpagetitle.Text = insertpagetitle
            End If
        End If
        ToolTip1.SetToolTip(linsertpagetitle, insertpagetitle)
        If Not insertpagetitle = Nothing Then
            step04.Enabled = True
            step04.Focus()
        End If
        get_input.Dispose()
    End Sub

    Private Sub step03_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step03.Click
        Step03_Code()
    End Sub

    Private Sub Step04_Code()
        Dim get_input As User_Input_Prompt = New User_Input_Prompt("Please enter the object description to used on the page. For example, if the database table deals with users, then the description could be ""User Details""")
        If get_input.ShowDialog() = DialogResult.OK Then
            If Not get_input.input.Text = Nothing Then
                insertobjectdescription = get_input.input.Text
                linsertobjectdescription.Text = insertobjectdescription
            End If
        End If
        ToolTip1.SetToolTip(linsertobjectdescription, insertobjectdescription)
        If Not insertobjectdescription = Nothing Then
            step05.Enabled = True
            step05.Focus()
        End If
        get_input.Dispose()
    End Sub

    Private Sub step04_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step04.Click
        Step04_Code()
    End Sub

    Private Sub Step05_Code()
        DialogResult = SelectDatabase.ShowDialog()
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then
            insertdatabase = SelectDatabase.FileName
            linsertdatabase.Text = insertdatabase
        End If
        ToolTip1.SetToolTip(linsertdatabase, insertdatabase)
        If Not insertdatabase = Nothing Then
            step06.Enabled = True
            step06.Focus()
        End If
    End Sub

    Private Sub step05_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step05.Click
        Step05_Code()
    End Sub

    Private Sub Step06_Code()
        Try
            Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & insertdatabase & ";")
            Conn.Open()
            Dim schemaTable As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            Dim frm_Select_Table_Dialog As New Select_Table_Dialog
            frm_Select_Table_Dialog.Activate()
            frm_Select_Table_Dialog.TableChoice = schemaTable

            If frm_Select_Table_Dialog.ShowDialog() = DialogResult.OK Then
                If Not frm_Select_Table_Dialog.Selected_Table.SelectedItem = Nothing Then
                    inserttable = frm_Select_Table_Dialog.Selected_Table.SelectedItem.ToString
                    linserttable.Text = inserttable
                End If
            End If
            ToolTip1.SetToolTip(linserttable, inserttable)
            If Not inserttable = Nothing Then
                step07.Enabled = True
                step07.Focus()
            End If
            frm_Select_Table_Dialog.Dispose()
            Conn.Close()
        Catch dberror As OleDb.OleDbException
            MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)

        Catch othererror As Exception
            MsgBox("Error encountered" & vbCrLf & othererror.Message)
        End Try
    End Sub

    Private Sub step06_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step06.Click
        Step06_Code()
    End Sub

    Private Sub Step07_Code()
        Try
            insertdatabase = SelectDatabase.FileName
            Dim f As FileInfo = New FileInfo(insertdatabase)
            insertdatabase = f.Name()

            ' Selected_Database = SelectDatabase.FileName

            Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & insertdatabase & ";")
            Conn.Open()
            Dim schemaTable As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            Dim frm_Select_Table_Dialog2 As New Select_Column_Dialog
            Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, inserttable, Nothing})
            frm_Select_Table_Dialog2.Activate()
            frm_Select_Table_Dialog2.TableChoice = schemaTable2




            If frm_Select_Table_Dialog2.ShowDialog() = DialogResult.OK Then
                If Not frm_Select_Table_Dialog2.Selected_Table.SelectedItem = Nothing Then
                    insertfieldkey = frm_Select_Table_Dialog2.Selected_Table.SelectedItem.ToString
                    linsertfieldkey.Text = insertfieldkey
                End If
            End If
            ToolTip1.SetToolTip(linsertfieldkey, insertfieldkey)
            If Not insertfieldkey = Nothing Then
                step09.Enabled = True
                step09.Focus()
            End If

            frm_Select_Table_Dialog2.Dispose()
            Conn.Close()

        Catch dberror As OleDb.OleDbException
            MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)

        Catch othererror As Exception
            MsgBox("Error encountered" & vbCrLf & othererror.Message)
        End Try
    End Sub

    Private Sub step07_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step07.Click
        Step07_Code()
    End Sub

    Private Sub Step08_Code()
        Try
            Dim filename As String = SelectDatabase.FileName

            Dim column_input_type = New Column_Input_Type_Mapping_Dialog
            column_input_type.ShowDialog()


            'If frm_Select_Table_Dialog2.ShowDialog() = DialogResult.OK Then
            '    If Not frm_Select_Table_Dialog2.Selected_Table.SelectedItem = Nothing Then
            '        insertfieldkey = frm_Select_Table_Dialog2.Selected_Table.SelectedItem.ToString
            '        linsertfieldkey.Text = insertfieldkey
            '    End If
            'End If
            'ToolTip1.SetToolTip(linsertfieldkey, insertfieldkey)
            'If Not insertfieldkey = Nothing Then
            '    step08.Enabled = True
            '    step08.Focus()
            'End If

            'frm_Select_Table_Dialog2.Dispose()
            'Conn.Close()

        Catch dberror As OleDb.OleDbException
            MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)

        Catch othererror As Exception
            MsgBox("Error encountered" & vbCrLf & othererror.Message)
        End Try
        step09.Enabled = True
        step09.Focus()
    End Sub

    Private Sub step08_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Step08_Code()
    End Sub

    Private Sub Step09_Code()


        Try

            Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & insertdatabase & ";")
            Conn.Open()

            Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, inserttable, Nothing})

            Dim myRow2 As DataRow
            Dim myCol2 As DataColumn
            Dim MyDataBaseColumn = New System.Collections.ArrayList

            Dim columnname As String
            Dim datatype As OleDb.OleDbType


            For Each myRow2 In schemaTable2.Rows
                columnname = ""


                For Each myCol2 In schemaTable2.Columns
                    If myCol2.ColumnName = "DATA_TYPE" Then
                        datatype = myRow2(myCol2)
                    End If
                    If myCol2.ColumnName = "COLUMN_NAME" Then
                        columnname = myRow2(myCol2).ToString()
                    End If

                Next
                If Not columnname.Equals("") Then
                    Dim dbcolumn As DatabaseColumn
                    dbcolumn.column_name = columnname
                    dbcolumn.data_type = datatype
                    MyDataBaseColumn.Add(dbcolumn)
                End If
            Next

            Conn.Close()



            Dim filereader As New StreamReader(template, True)
            Dim lineread As String = filereader.ReadLine
            Dim filewriter As New StreamWriter(insertfilename_fullpath, False, System.Text.Encoding.ASCII)
            Dim linewritten As Boolean
            Dim size As Integer
            Do While Not lineread Is Nothing
                lineread = lineread.Trim()
                linewritten = False
                lineread = lineread.Replace("#insertpagetitle#", insertpagetitle)
                lineread = lineread.Replace("#insertfilename#", insertfilename)
                lineread = lineread.Replace("#insertdatabase#", insertdatabase)
                lineread = lineread.Replace("#insertfieldkey#", insertfieldkey)
                lineread = lineread.Replace("#inserttable#", inserttable)
                lineread = lineread.Replace("#insertobjectdescription#", insertobjectdescription)
                If lineread.StartsWith("#retrievevariables#") Then
                    size = 0
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        filewriter.WriteLine(dbcolumn.column_name & " = trim(replace(request.form(""" & dbcolumn.column_name & """),""'"",""`""))")
                        If dbcolumn.data_type.ToString().ToLower = "boolean" Then
                            filewriter.WriteLine("If " & dbcolumn.column_name & " = ""ON"" Then")
                            filewriter.WriteLine(dbcolumn.column_name & " = 1")
                            filewriter.WriteLine("Else")
                            filewriter.WriteLine(dbcolumn.column_name & " = 0")
                            filewriter.WriteLine("End If")
                        End If
                        If dbcolumn.data_type.ToString().ToLower = "integer" Then
                            filewriter.WriteLine("If " & dbcolumn.column_name & " = """" Then")
                            filewriter.WriteLine(dbcolumn.column_name & " = ""0""")
                            filewriter.WriteLine("End If")
                        End If
                        size = size + 1
                    End While
                    linewritten = True
                End If
                If lineread.StartsWith("#createtextboxes#") Then
                    size = 0
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        filewriter.WriteLine("<tr>")
                        filewriter.WriteLine("<td width=""50%"">" & dbcolumn.column_name & ":</td>")
                        If dbcolumn.data_type.ToString().ToLower = "boolean" Then
                            filewriter.WriteLine("<td width=""50%""> <input type=""checkbox"" name=""" & dbcolumn.column_name & """ value=""ON""></td>")
                        Else
                            filewriter.WriteLine("<td width=""50%""> <input type=""text"" name=""" & dbcolumn.column_name & """ size=""20""></td>")
                        End If
                        filewriter.WriteLine("</tr>")
                        size = size + 1
                    End While
                    linewritten = True
                End If

                If lineread.StartsWith("#createfilledtextboxes#") Then
                    size = 0
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        filewriter.WriteLine("<tr>")
                        filewriter.WriteLine("<td width=""50%"">" & dbcolumn.column_name & ":</td>")
                        filewriter.WriteLine("<td width=""50%""> <input type=""text"" name=""" & dbcolumn.column_name & """ size=""20"" value=""<% response.write objRec.fields(""" & dbcolumn.column_name & """).value %>""></td>")
                        filewriter.WriteLine("</tr>")
                        size = size + 1
                    End While
                    linewritten = True
                End If

                If lineread.StartsWith("#createdisplaytable#") Then
                    size = 0

                    filewriter.WriteLine("<tr>")
                    filewriter.WriteLine("<td align=""center"" valign=""top""><b>" & insertfieldkey & "</b></td>")
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        If Not insertfieldkey = dbcolumn.column_name Then
                            filewriter.WriteLine("<td align=""center"" valign=""top""><b>" & dbcolumn.column_name & "</b></td>")
                        End If
                        size = size + 1
                    End While
                    filewriter.WriteLine("</tr>")
                    filewriter.WriteLine("<% while not objRec.eof %>")
                    filewriter.WriteLine("<tr>")
                    size = 0
                    filewriter.WriteLine("<td valign=""top""><% response.write objRec.Fields(""" & insertfieldkey & """).value %></td>")
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        If Not insertfieldkey = dbcolumn.column_name Then
                            filewriter.WriteLine("<td valign=""top""><% response.write objRec.Fields(""" & dbcolumn.column_name & """).value %></td>")
                        End If
                        size = size + 1
                    End While
                    filewriter.WriteLine("</tr>")
                    filewriter.WriteLine("<%" & vbCrLf & "objRec.MoveNext")
                    filewriter.WriteLine("wend" & vbCrLf & "%>")
                    linewritten = True
                End If


                If lineread.StartsWith("#validateentries#") Then

                    Dim linetowrite As String
                    linetowrite = "if "
                    size = 0
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        linetowrite = linetowrite & dbcolumn.column_name & " = """" or "
                        size = size + 1
                    End While
                    linetowrite = linetowrite.Remove(linetowrite.Length - 3, 3) & "then"
                    filewriter.WriteLine(linetowrite)
                    filewriter.WriteLine("insertflag = ""false""")
                    filewriter.WriteLine("%>")
                    filewriter.WriteLine("<script Language=""javascript"">")
                    filewriter.WriteLine("alert(""The " & insertobjectdescription & " cannot be added as not all of the required text fields have been filled in"");")
                    filewriter.WriteLine("window.location = """ & insertfilename & """")
                    filewriter.WriteLine("</script>")
                    filewriter.WriteLine("<%")
                    filewriter.WriteLine("End If")
                    linewritten = True
                End If

                If lineread.StartsWith("#generateupdateSQL#") Then
                    size = 0
                    Dim linetowrite As String
                    linetowrite = "countersql = ""update " & inserttable & " set "
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        If dbcolumn.data_type.ToString().ToLower = "wchar" Then
                            linetowrite = linetowrite & dbcolumn.column_name & " = '"" & " & dbcolumn.column_name & " & ""', "
                        Else
                            linetowrite = linetowrite & dbcolumn.column_name & " = "" & " & dbcolumn.column_name & " & "", "
                        End If
                        size = size + 1
                    End While
                    linetowrite = linetowrite.Remove(linetowrite.Length - 2, 2)
                    linetowrite = linetowrite & " where " & insertfieldkey & " = '"" & old" & insertfieldkey & " & ""'"""
                    filewriter.WriteLine(linetowrite)
                    linewritten = True
                End If

                If lineread.StartsWith("#generateinsertSQL#") Then
                    size = 0
                    Dim linetowrite As String
                    linetowrite = "countersql = ""insert into " & inserttable & "("
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        linetowrite = linetowrite & dbcolumn.column_name & ", "
                        size = size + 1
                    End While
                    linetowrite = linetowrite.Remove(linetowrite.Length - 2, 2)
                    linetowrite = linetowrite & ") values ("
                    size = 0
                    While size < MyDataBaseColumn.Count()
                        Dim dbcolumn As DatabaseColumn = MyDataBaseColumn.Item(size)
                        If dbcolumn.data_type.ToString().ToLower = "wchar" Then
                            linetowrite = linetowrite & "'"" & " & dbcolumn.column_name & " & ""', "
                        Else
                            linetowrite = linetowrite & """ & " & dbcolumn.column_name & " & "", "
                        End If
                        size = size + 1
                    End While
                    linetowrite = linetowrite.Remove(linetowrite.Length - 2, 2)
                    linetowrite = linetowrite & ")"""
                    filewriter.WriteLine(linetowrite)
                    linewritten = True
                End If

                If linewritten = False Then
                    filewriter.WriteLine(lineread)
                End If
                lineread = filereader.ReadLine
            Loop

            filereader.Close()
            filewriter.Close()
            MsgBox("Process Completed")
        Catch dberror As OleDb.OleDbException
            MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)

        Catch othererror As Exception
            MsgBox("Error encountered" & vbCrLf & othererror.Message)
        End Try

    End Sub

    Private Sub step09_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles step09.Click
        Step09_Code()
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Dim About_Screen1 As About_Screen = New About_Screen
        DialogResult = About_Screen1.ShowDialog()
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DialogResult = LoadJobFile.ShowDialog()
        Dim jobfilename As String
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then
            jobfilename = LoadJobFile.FileName
            Dim filereader As New StreamReader(jobfilename, True)
            Dim lineread As String = filereader.ReadLine

            Do While Not lineread Is Nothing
                lineread = lineread.Trim()
                If lineread.StartsWith("Template: ") Then
                    template = lineread.Remove(0, 10)
                    ltemplate.Text = lineread.Remove(0, 10)
                    SelectTemplate.FileName = template
                    ToolTip1.SetToolTip(ltemplate, template)
                    step01.Enabled = True
                    step02.Enabled = True
                End If
                If lineread.StartsWith("ASP Filename: ") Then
                    insertfilename = lineread.Remove(0, 14)
                    insertfilename = insertfilename.Remove(0, insertfilename.LastIndexOf("\") + 1)
                    insertfilename_fullpath = lineread.Remove(0, 14)
                    linsertfilename.Text = insertfilename
                    SaveResult.FileName = insertfilename_fullpath
                    ToolTip1.SetToolTip(linsertfilename, insertfilename_fullpath)
                    step02.Enabled = True
                    step03.Enabled = True
                End If
                If lineread.StartsWith("Page Title: ") Then
                    insertpagetitle = lineread.Remove(0, 12)
                    linsertpagetitle.Text = lineread.Remove(0, 12)
                    ToolTip1.SetToolTip(linsertpagetitle, insertpagetitle)
                    step03.Enabled = True
                    step04.Enabled = True
                End If
                If lineread.StartsWith("Object Description: ") Then
                    insertobjectdescription = lineread.Remove(0, 20)
                    linsertobjectdescription.Text = lineread.Remove(0, 20)
                    ToolTip1.SetToolTip(linsertobjectdescription, insertobjectdescription)
                    step04.Enabled = True
                    step05.Enabled = True
                End If
                If lineread.StartsWith("Database: ") Then
                    insertdatabase = lineread.Remove(0, 10)
                    linsertdatabase.Text = lineread.Remove(0, 10)
                    SelectDatabase.FileName = insertdatabase
                    ToolTip1.SetToolTip(linsertdatabase, insertdatabase)
                    step05.Enabled = True
                    step06.Enabled = True
                End If
                If lineread.StartsWith("Table: ") Then
                    inserttable = lineread.Remove(0, 7)
                    linserttable.Text = lineread.Remove(0, 7)
                    ToolTip1.SetToolTip(linserttable, inserttable)
                    step06.Enabled = True
                    step07.Enabled = True
                End If
                If lineread.StartsWith("Primary Key: ") Then
                    insertfieldkey = lineread.Remove(0, 13)
                    linsertfieldkey.Text = lineread.Remove(0, 13)
                    ToolTip1.SetToolTip(linsertfieldkey, insertfieldkey)
                    step07.Enabled = True
                    'step08.Enabled = True
                End If

                lineread = filereader.ReadLine
            Loop
            filereader.Close()
        End If
    End Sub

End Class