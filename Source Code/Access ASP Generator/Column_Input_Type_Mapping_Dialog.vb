Imports System
Imports System.IO
Imports System.Collections


Public Class Column_Input_Type_Mapping_Dialog
    Inherits System.Windows.Forms.Form

    Dim MyDataBaseColumn As String

    Public Structure DatabaseColumn
        Public column_name As String
        Public data_type As OleDb.OleDbType
        Public input_type As String
        Public display_column As Boolean
    End Structure

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
    Friend WithEvents Column_Panel As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Column_Input_Type_Mapping_Dialog))
        Me.Column_Panel = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Column_Panel.SuspendLayout()
        Me.SuspendLayout()
        '
        'Column_Panel
        '
        Me.Column_Panel.Controls.Add(Me.Label1)
        Me.Column_Panel.Location = New System.Drawing.Point(16, 16)
        Me.Column_Panel.Name = "Column_Panel"
        Me.Column_Panel.Size = New System.Drawing.Size(360, 232)
        Me.Column_Panel.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'Column_Input_Type_Mapping_Dialog
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(392, 294)
        Me.Controls.Add(Me.Column_Panel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Column_Input_Type_Mapping_Dialog"
        Me.Text = "Column_Input_Type_Mapping_Dialog"
        Me.Column_Panel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Column_Input_Type_Mapping_Dialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'insertdatabase = SelectDatabase.FileName
        'Dim f As FileInfo = New FileInfo(insertdatabase)
        'insertdatabase = f.Name()

        'Selected_Database = SelectDatabase.FileName

        'Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & insertdatabase & ";")
        'Conn.Open()
        'Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, inserttable, Nothing})

        'Dim myRow2 As DataRow
        'Dim myCol2 As DataColumn
        'Dim MyDataBaseColumn = New System.Collections.ArrayList

        'Dim columnname As String
        'Dim datatype As OleDb.OleDbType


        'For Each myRow2 In schemaTable2.Rows
        '    columnname = ""


        '    For Each myCol2 In schemaTable2.Columns
        '        If myCol2.ColumnName = "DATA_TYPE" Then
        '            datatype = myRow2(myCol2)
        '        End If
        '        If myCol2.ColumnName = "COLUMN_NAME" Then
        '            columnname = myRow2(myCol2).ToString()
        '        End If

        '    Next
        '    If Not columnname.Equals("") Then
        '        Dim dbcolumn As DatabaseColumn
        '        dbcolumn.column_name = columnname
        '        dbcolumn.data_type = datatype
        '        MyDataBaseColumn.Add(dbcolumn)
        '    End If
        'Next
    End Sub
End Class
