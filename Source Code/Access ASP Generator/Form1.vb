Imports System
Imports System.IO
Imports System.Collections

Public Class Main_Program
    Inherits System.Windows.Forms.Form

    '*****************************************'
    ' New variables
    Dim insertdatabase As String
    Dim insertfieldkey As String
    Dim inserttable As String
    Dim insertobjectdescription As String
    Dim insertfilename As String
    Dim template As String
    '*****************************************'

    Dim inputtedpath As String
    Dim recordcounter As Integer
    Dim selecteddrive As String
    Dim selectedpath As String
    Private Selected_Database As String
    Private Selected_Table As String

    Dim filenamelist As New System.Collections.ArrayList()

    Dim checklistset As Integer
    Dim maxchecklistset As Integer

    Public Structure DatabaseColumn
        Public column_name As String
        Public data_type As OleDb.OleDbType
    End Structure

    Public Structure MyFileItem
        Public FilenameString As String
        Public SizeString As String
        Public CheckedStatus As Boolean
    End Structure

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        selectedpath = ""
        selecteddrive = ""
        recordcounter = 0
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
    Friend WithEvents Select_Database_Dialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Datasource As System.Windows.Forms.Label
    Friend WithEvents Columns As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents linsertdatabase As System.Windows.Forms.Label
    Friend WithEvents linserttable As System.Windows.Forms.Label
    Friend WithEvents linsertfieldkey As System.Windows.Forms.Label
    Friend WithEvents linsertobjectdescription As System.Windows.Forms.Label
    Friend WithEvents linsertfilename As System.Windows.Forms.Label
    Friend WithEvents insertvalidate As System.Windows.Forms.CheckBox
    Friend WithEvents SelectTemplate As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveResult As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ltemplate As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Select_Database_Dialog = New System.Windows.Forms.OpenFileDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Columns = New System.Windows.Forms.Label()
        Me.Datasource = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.linsertdatabase = New System.Windows.Forms.Label()
        Me.linserttable = New System.Windows.Forms.Label()
        Me.linsertobjectdescription = New System.Windows.Forms.Label()
        Me.linsertfilename = New System.Windows.Forms.Label()
        Me.linsertfieldkey = New System.Windows.Forms.Label()
        Me.insertvalidate = New System.Windows.Forms.CheckBox()
        Me.SelectTemplate = New System.Windows.Forms.OpenFileDialog()
        Me.SaveResult = New System.Windows.Forms.SaveFileDialog()
        Me.ltemplate = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Gainsboro
        Me.Button2.Location = New System.Drawing.Point(24, 40)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Select Database"
        Me.ToolTip1.SetToolTip(Me.Button2, "Allows you to select a database to which the results will be sent to")
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Panel1.Location = New System.Drawing.Point(168, 16)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 144)
        Me.Panel1.TabIndex = 9
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Gainsboro
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(24, 32)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 23)
        Me.Button3.TabIndex = 8
        Me.Button3.Text = "Update Database"
        Me.ToolTip1.SetToolTip(Me.Button3, "Allows you to update the selected database")
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Columns, Me.Panel1, Me.Button2, Me.Datasource})
        Me.GroupBox2.Location = New System.Drawing.Point(24, 352)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(688, 192)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Step 2 - Select a Data Source"
        '
        'Columns
        '
        Me.Columns.Location = New System.Drawing.Point(584, 168)
        Me.Columns.Name = "Columns"
        Me.Columns.Size = New System.Drawing.Size(88, 16)
        Me.Columns.TabIndex = 15
        Me.Columns.Text = "0 Columns"
        Me.Columns.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Datasource
        '
        Me.Datasource.Location = New System.Drawing.Point(16, 168)
        Me.Datasource.Name = "Datasource"
        Me.Datasource.Size = New System.Drawing.Size(568, 16)
        Me.Datasource.TabIndex = 14
        Me.Datasource.Text = "Data Source: "
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button3})
        Me.GroupBox3.Location = New System.Drawing.Point(24, 552)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(688, 144)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Step 3 - Update the Database"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(24, 56)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 15
        Me.Button1.Text = "Button1"
        '
        'linsertdatabase
        '
        Me.linsertdatabase.Location = New System.Drawing.Point(24, 168)
        Me.linsertdatabase.Name = "linsertdatabase"
        Me.linsertdatabase.Size = New System.Drawing.Size(672, 16)
        Me.linsertdatabase.TabIndex = 16
        Me.linsertdatabase.Text = "Database:"
        '
        'linserttable
        '
        Me.linserttable.Location = New System.Drawing.Point(24, 192)
        Me.linserttable.Name = "linserttable"
        Me.linserttable.Size = New System.Drawing.Size(672, 16)
        Me.linserttable.TabIndex = 18
        Me.linserttable.Text = "Table:"
        '
        'linsertobjectdescription
        '
        Me.linsertobjectdescription.Location = New System.Drawing.Point(24, 120)
        Me.linsertobjectdescription.Name = "linsertobjectdescription"
        Me.linsertobjectdescription.Size = New System.Drawing.Size(672, 16)
        Me.linsertobjectdescription.TabIndex = 19
        Me.linsertobjectdescription.Text = "Object Description: "
        '
        'linsertfilename
        '
        Me.linsertfilename.Location = New System.Drawing.Point(24, 144)
        Me.linsertfilename.Name = "linsertfilename"
        Me.linsertfilename.Size = New System.Drawing.Size(672, 16)
        Me.linsertfilename.TabIndex = 20
        Me.linsertfilename.Text = "Filename: "
        '
        'linsertfieldkey
        '
        Me.linsertfieldkey.Location = New System.Drawing.Point(24, 216)
        Me.linsertfieldkey.Name = "linsertfieldkey"
        Me.linsertfieldkey.Size = New System.Drawing.Size(672, 16)
        Me.linsertfieldkey.TabIndex = 22
        Me.linsertfieldkey.Text = "Primary Key:"
        '
        'insertvalidate
        '
        Me.insertvalidate.Checked = True
        Me.insertvalidate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.insertvalidate.Location = New System.Drawing.Point(24, 240)
        Me.insertvalidate.Name = "insertvalidate"
        Me.insertvalidate.TabIndex = 23
        Me.insertvalidate.Text = "Validation?"
        '
        'SaveResult
        '
        Me.SaveResult.FileName = "doc1"
        '
        'ltemplate
        '
        Me.ltemplate.Location = New System.Drawing.Point(24, 96)
        Me.ltemplate.Name = "ltemplate"
        Me.ltemplate.Size = New System.Drawing.Size(672, 16)
        Me.ltemplate.TabIndex = 24
        Me.ltemplate.Text = "Template:"
        '
        'Main_Program
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(736, 703)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ltemplate, Me.insertvalidate, Me.linsertfieldkey, Me.linsertfilename, Me.linsertobjectdescription, Me.linserttable, Me.linsertdatabase, Me.Button1, Me.GroupBox2, Me.GroupBox3})
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(744, 730)
        Me.MinimumSize = New System.Drawing.Size(744, 730)
        Me.Name = "Main_Program"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manga CD List Generator"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        DialogResult = Select_Database_Dialog.ShowDialog()
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then

            Try
                Selected_Database = Select_Database_Dialog.FileName


                Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Selected_Database & ";")
                Conn.Open()
                Dim schemaTable As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

                Dim frm_Select_Table_Dialog As New Select_Table_Dialog()
                frm_Select_Table_Dialog.Activate()
                frm_Select_Table_Dialog.TableChoice = schemaTable
                frm_Select_Table_Dialog.ShowDialog()
                Dim tableresult As String = frm_Select_Table_Dialog.Selected_Table.SelectedItem.ToString
                Selected_Table = tableresult
                frm_Select_Table_Dialog.Dispose()
                Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tableresult, Nothing})

                Dim myRow2 As DataRow
                Dim myCol2 As DataColumn

                Panel1.Controls.Clear()

                '    RichTextBox1.Clear()

                Dim ordinal As Integer
                Dim columnname As String
                Dim datatype As OleDb.OleDbType
                For Each myRow2 In schemaTable2.Rows
                    ordinal = 0
                    columnname = ""


                    For Each myCol2 In schemaTable2.Columns
                        If myCol2.ColumnName = "DATA_TYPE" Then
                            datatype = myRow2(myCol2)
                        End If
                        If myCol2.ColumnName = "COLUMN_NAME" Then
                            columnname = myRow2(myCol2).ToString()
                        End If
                        If myCol2.ColumnName = "ORDINAL_POSITION" Then
                            ordinal = Val(myRow2(myCol2).ToString())
                        End If
                    Next
                    ordinal = ordinal - 1
                    If Not columnname.Equals("") Then
                        'MsgBox(myRow2(myCol2).ToString())
                        Dim LabelMiniMe As New System.Windows.Forms.Label()
                        LabelMiniMe.Location = New System.Drawing.Point(0, (ordinal * 24))
                        LabelMiniMe.Name = "Label_" & columnname
                        LabelMiniMe.Size = New System.Drawing.Size(136, 16)
                        LabelMiniMe.Text = columnname
                        Panel1.Controls.Add(LabelMiniMe)
                        LabelMiniMe.Visible = True

                        Dim ComboBoxMiniMe As New System.Windows.Forms.ComboBox()
                        ComboBoxMiniMe.Location = New System.Drawing.Point(140, (ordinal * 24))
                        ComboBoxMiniMe.Name = "ComboBox_" & columnname
                        ComboBoxMiniMe.Size = New System.Drawing.Size(136, 16)
                        If datatype.ToString().ToLower = "wchar" Then
                            ComboBoxMiniMe.Text = "'" & columnname & "'"
                            ComboBoxMiniMe.Items.Add("Selected Item String")
                            '  If Display_Size.Checked = True Then
                            ComboBoxMiniMe.Items.Add("File Size String")
                            ' End If

                        Else
                            ComboBoxMiniMe.Text = columnname

                        End If
                        ComboBoxMiniMe.Items.Add("Ignore Column")
                        Panel1.Controls.Add(ComboBoxMiniMe)
                        ComboBoxMiniMe.Visible = True
                        If datatype.ToString().ToLower = "boolean" Then
                            ComboBoxMiniMe.Items.Add(1)
                            ComboBoxMiniMe.Items.Add(0)
                            ComboBoxMiniMe.SelectedIndex = 0
                            ComboBoxMiniMe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
                        End If

                        Dim LabelMiniM2 As New System.Windows.Forms.Label()
                        LabelMiniM2.Location = New System.Drawing.Point(280, (ordinal * 24))
                        LabelMiniM2.Name = "Label_TYPE_" & columnname
                        LabelMiniM2.Size = New System.Drawing.Size(50, 16)
                        LabelMiniM2.Text = datatype.ToString()
                        Panel1.Controls.Add(LabelMiniM2)
                        LabelMiniM2.Visible = True

                    End If
                Next
                Conn.Close()
                Button3.Enabled = True
                Datasource.Text = "Data Source: " & Selected_Database & "      Table: " & Selected_Table
                Columns.Text = (Panel1.Controls.Count / 3) & " Columns"
            Catch dberror As OleDb.OleDbException
                MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)

            Catch othererror As Exception
                MsgBox("Error encountered" & vbCrLf & othererror.Message)
            End Try
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Selected_Database & ";")
            Conn.Open()
            Dim counter As Integer
            Dim itemtoadd As MyFileItem
            Dim stoppedflag As Boolean = False
            For counter = 0 To filenamelist.Count - 1
                itemtoadd = filenamelist.Item(counter)
                If itemtoadd.CheckedStatus = True Then
                    Try
                        Dim sqlstr As String
                        sqlstr = "insert into " & Selected_Table & "("
                        Dim runne As Integer
                        For runne = 0 To Panel1.Controls.Count - 1 Step 3
                            If Not Panel1.Controls(runne + 1).Text = "Ignore Column" Then
                                sqlstr = sqlstr & Panel1.Controls(runne).Text & ","
                            End If
                        Next
                        sqlstr = sqlstr.Remove(sqlstr.Length - 1, 1)

                        sqlstr = sqlstr & ") values ("
                        Dim runner As Integer
                        For runner = 1 To Panel1.Controls.Count - 1 Step 3
                            If Panel1.Controls(runner).Text = "Selected Item String" Then
                                'Dim itemtoadd As MyFileItem
                                'itemtoadd = filenamelist.Item(counter)
                                If itemtoadd.CheckedStatus = True Then
                                    sqlstr = sqlstr & "'" & itemtoadd.FilenameString.Replace("'", "`") & "',"
                                End If
                            Else
                                If Panel1.Controls(runner).Text = "File Size String" Then
                                    'Dim itemtoadd As MyFileItem
                                    'itemtoadd = filenamelist.Item(counter)
                                    If itemtoadd.CheckedStatus = True Then
                                        sqlstr = sqlstr & "'" & itemtoadd.SizeString & "',"
                                    End If

                                Else
                                    If Not Panel1.Controls(runner).Text = "Ignore Column" Then
                                        sqlstr = sqlstr & Panel1.Controls(runner).Text & ","
                                    End If
                                End If
                            End If
                        Next
                        sqlstr = sqlstr.Remove(sqlstr.Length - 1, 1)
                        sqlstr = sqlstr & ")"
                        '  RichTextBox1.Text = sqlstr & vbCrLf & RichTextBox1.Text

                        Dim recset As Data.OleDb.OleDbCommand = New Data.OleDb.OleDbCommand(sqlstr, Conn)
                        recset.ExecuteNonQuery()
                    Catch sqlerror As Exception
                        Dim answer As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Cannot update the datbase" & vbCrLf & sqlerror.Message, MsgBoxStyle.OKCancel)
                        If answer = MsgBoxResult.Abort Or answer = MsgBoxResult.Cancel Then
                            ' MsgBox("Stopping further updates to the Database")
                            ' updateresultlabel.Text = "Stopping further updates to the Database"
                            stoppedflag = True
                            Exit For
                        End If
                    End Try
                End If
            Next
            '  RichTextBox1.Text = Now() & vbCrLf & RichTextBox1.Text

            Conn.Close()
            If stoppedflag = False Then
                'MsgBox("Database Successfully Updated")
                '    updateresultlabel.Text = "Database Successfully Updated"
            End If
        Catch dberror As OleDb.OleDbException
            'MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)
            '  updateresultlabel.Text = "Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message
        Catch othererror As Exception
            'MsgBox("Error encountered" & vbCrLf & othererror.Message)
            ' updateresultlabel.Text = "Error encountered" & vbCrLf & othererror.Message
        End Try

    End Sub



    Private Sub Main_Program_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        DialogResult = SelectTemplate.ShowDialog()
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then
            template = SelectTemplate.FileName
            ltemplate.Text = "Filename: " & template
        End If
        insertobjectdescription = InputBox("Enter Object Description", "Object Description", "")
        linsertobjectdescription.Text = "Object Description: " & insertobjectdescription
        DialogResult = SaveResult.ShowDialog()
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then
            insertfilename = SaveResult.FileName
            insertfilename = insertfilename.Remove(0, insertfilename.LastIndexOf("\") + 1)
            linsertfilename.Text = "Filename: " & insertfilename
        End If
        DialogResult = Select_Database_Dialog.ShowDialog()
        If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then

            Try
                insertdatabase = Select_Database_Dialog.FileName
                linsertdatabase.Text = "Database: " & insertdatabase
                Selected_Database = Select_Database_Dialog.FileName


                Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Selected_Database & ";")
                Conn.Open()
                Dim schemaTable As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

                Dim frm_Select_Table_Dialog As New Select_Table_Dialog()
                frm_Select_Table_Dialog.Activate()
                frm_Select_Table_Dialog.TableChoice = schemaTable
                frm_Select_Table_Dialog.ShowDialog()
                Dim tableresult As String = frm_Select_Table_Dialog.Selected_Table.SelectedItem.ToString
                inserttable = tableresult
                linserttable.Text = "Table: " & inserttable
                Selected_Table = tableresult
                frm_Select_Table_Dialog.Dispose()

                Dim frm_Select_Table_Dialog2 As New Select_Column_Dialog()
                Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tableresult, Nothing})
                frm_Select_Table_Dialog2.Activate()
                frm_Select_Table_Dialog2.TableChoice = schemaTable2
                frm_Select_Table_Dialog2.ShowDialog()
                tableresult = frm_Select_Table_Dialog2.Selected_Table.SelectedItem.ToString
                frm_Select_Table_Dialog2.Dispose()
                insertfieldkey = tableresult
                linsertfieldkey.Text = "Primary Key: " & insertfieldkey

                Dim myRow2 As DataRow
                Dim myCol2 As DataColumn
                Dim MyDataBaseColumn = New System.Collections.ArrayList()

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
                Dim filewriter As New StreamWriter(SaveResult.FileName, False, System.Text.Encoding.ASCII)
                Dim linewritten As Boolean
                Dim size As Integer
                Do While Not lineread Is Nothing
                    linewritten = False
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

                    If lineread.StartsWith("#validateentries#") Then
                        If insertvalidate.Checked = True Then
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
                        End If
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



        End If

    End Sub

End Class