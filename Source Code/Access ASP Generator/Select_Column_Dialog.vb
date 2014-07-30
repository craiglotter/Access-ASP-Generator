Public Class Select_Column_Dialog
    Inherits System.Windows.Forms.Form

    Private TableChoiceValue As DataTable

    Property TableChoice() As DataTable
        Get
            Return TableChoiceValue
        End Get
        Set(ByVal Value As DataTable)
            TableChoiceValue = Value
        End Set
    End Property

#Region " Windows Form Designer generated code "

    'Public Sub New()
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()

    '    'Add any initialization after the InitializeComponent() call
    '    Type = 1
    'End Sub

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
    Friend WithEvents Selected_Table As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents accept_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Select_Column_Dialog))
        Me.accept_button = New System.Windows.Forms.Button
        Me.Selected_Table = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cancel_button = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'accept_button
        '
        Me.accept_button.BackColor = System.Drawing.SystemColors.Control
        Me.accept_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.accept_button.Location = New System.Drawing.Point(240, 40)
        Me.accept_button.Name = "accept_button"
        Me.accept_button.TabIndex = 0
        Me.accept_button.Text = "Accept"
        '
        'Selected_Table
        '
        Me.Selected_Table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Selected_Table.Location = New System.Drawing.Point(16, 40)
        Me.Selected_Table.MaxDropDownItems = 20
        Me.Selected_Table.Name = "Selected_Table"
        Me.Selected_Table.Size = New System.Drawing.Size(216, 21)
        Me.Selected_Table.Sorted = True
        Me.Selected_Table.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(192, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Please select a table column."
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(320, 40)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 3
        Me.cancel_button.Text = "Cancel"
        '
        'Select_Column_Dialog
        '
        Me.AcceptButton = Me.accept_button
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cancel_button
        Me.ClientSize = New System.Drawing.Size(416, 78)
        Me.ControlBox = False
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Selected_Table)
        Me.Controls.Add(Me.accept_button)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Select_Column_Dialog"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select a Table Column"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Select_Column_Dialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With Selected_Table
            Dim myRow As DataRow
            Dim myCol As DataColumn

            For Each myRow In TableChoice.Rows
                For Each myCol In TableChoice.Columns
                    If myCol.ColumnName = "COLUMN_NAME" Then
                        .Items.Add(myRow(myCol).ToString())
                        .SelectedIndex = 0
                    End If
                Next
            Next

        End With
    End Sub

End Class
