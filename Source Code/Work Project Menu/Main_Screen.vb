Imports Microsoft.Win32
Imports System.IO


Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim projectdirectory As String = ""
    Dim runprog1 As String = ""
    Dim runprog2 As String = ""

    '  Dim WithEvents dt As Data.DataTable
    Public dataloaded As Boolean = False
    Private splash_loader As Splash_Screen

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        '    dt = New Data.DataTable("projects")
        '   dt.Columns.Add("Project Name")
        '   dt.Columns.Add("Date")
        '    dt.Columns.Add("Project Folder")
        '
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
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
    Friend WithEvents FullError As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents dv As System.Data.DataView
    Friend WithEvents txtBaseFolder As System.Windows.Forms.TextBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents ButtonFolderBrowse As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataSet1 As System.Data.DataSet
    Friend WithEvents DataTable1 As System.Data.DataTable
    Friend WithEvents DataColumn1 As System.Data.DataColumn
    Friend WithEvents DataColumn2 As System.Data.DataColumn
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents buildlabel As System.Windows.Forms.Label
    Friend WithEvents DataGridTextBoxColumn4 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataColumn4 As System.Data.DataColumn
    Friend WithEvents DataColumn3 As System.Data.DataColumn
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.FullError = New System.Windows.Forms.CheckBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtBaseFolder = New System.Windows.Forms.TextBox
        Me.ButtonFolderBrowse = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.dv = New System.Data.DataView
        Me.DataTable1 = New System.Data.DataTable
        Me.DataColumn1 = New System.Data.DataColumn
        Me.DataColumn2 = New System.Data.DataColumn
        Me.DataColumn4 = New System.Data.DataColumn
        Me.DataColumn3 = New System.Data.DataColumn
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn4 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.buildlabel = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label5 = New System.Windows.Forms.Label
        Me.DataSet1 = New System.Data.DataSet
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dv, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataTable1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FullError
        '
        Me.FullError.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FullError.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FullError.Location = New System.Drawing.Point(756, 32)
        Me.FullError.Name = "FullError"
        Me.FullError.Size = New System.Drawing.Size(12, 12)
        Me.FullError.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.FullError, "If Checked, Error Handling Routine enters Full Exception Mode")
        '
        'txtBaseFolder
        '
        Me.txtBaseFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBaseFolder.BackColor = System.Drawing.Color.White
        Me.txtBaseFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBaseFolder.ForeColor = System.Drawing.Color.Black
        Me.txtBaseFolder.Location = New System.Drawing.Point(432, 8)
        Me.txtBaseFolder.Name = "txtBaseFolder"
        Me.txtBaseFolder.ReadOnly = True
        Me.txtBaseFolder.Size = New System.Drawing.Size(272, 20)
        Me.txtBaseFolder.TabIndex = 4
        Me.txtBaseFolder.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtBaseFolder, "The base folder containing the various work projects")
        '
        'ButtonFolderBrowse
        '
        Me.ButtonFolderBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonFolderBrowse.BackColor = System.Drawing.Color.DarkSlateBlue
        Me.ButtonFolderBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonFolderBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonFolderBrowse.ForeColor = System.Drawing.Color.White
        Me.ButtonFolderBrowse.Location = New System.Drawing.Point(712, 8)
        Me.ButtonFolderBrowse.Name = "ButtonFolderBrowse"
        Me.ButtonFolderBrowse.Size = New System.Drawing.Size(56, 20)
        Me.ButtonFolderBrowse.TabIndex = 7
        Me.ButtonFolderBrowse.Text = "BROWSE"
        Me.ToolTip1.SetToolTip(Me.ButtonFolderBrowse, "Launches the Folder Browser Dialog")
        '
        'DataGrid1
        '
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.MediumSlateBlue
        Me.DataGrid1.CaptionForeColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionText = "Work Projects"
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.DataSource = Me.dv
        Me.DataGrid1.FlatMode = True
        Me.DataGrid1.ForeColor = System.Drawing.Color.DimGray
        Me.DataGrid1.GridLineColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.HeaderBackColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.Black
        Me.DataGrid1.LinkColor = System.Drawing.Color.Black
        Me.DataGrid1.Location = New System.Drawing.Point(0, 0)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsBackColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.ParentRowsForeColor = System.Drawing.Color.Black
        Me.DataGrid1.ParentRowsVisible = False
        Me.DataGrid1.PreferredColumnWidth = 100
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.RowHeadersVisible = False
        Me.DataGrid1.RowHeaderWidth = 30
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.LightSteelBlue
        Me.DataGrid1.SelectionForeColor = System.Drawing.Color.Black
        Me.DataGrid1.Size = New System.Drawing.Size(758, 166)
        Me.DataGrid1.TabIndex = 1
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        Me.ToolTip1.SetToolTip(Me.DataGrid1, "A list of valid work projects located in the base directory")
        '
        'dv
        '
        Me.dv.AllowDelete = False
        Me.dv.AllowEdit = False
        Me.dv.AllowNew = False
        Me.dv.ApplyDefaultSort = True
        Me.dv.Table = Me.DataTable1
        '
        'DataTable1
        '
        Me.DataTable1.Columns.AddRange(New System.Data.DataColumn() {Me.DataColumn1, Me.DataColumn2, Me.DataColumn4, Me.DataColumn3})
        Me.DataTable1.TableName = "Projects"
        '
        'DataColumn1
        '
        Me.DataColumn1.AllowDBNull = False
        Me.DataColumn1.ColumnName = "Project Name"
        '
        'DataColumn2
        '
        Me.DataColumn2.AllowDBNull = False
        Me.DataColumn2.ColumnName = "Conception Date"
        '
        'DataColumn4
        '
        Me.DataColumn4.AllowDBNull = False
        Me.DataColumn4.ColumnName = "Latest Build"
        '
        'DataColumn3
        '
        Me.DataColumn3.ColumnName = "Working Directory"
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.AlternatingBackColor = System.Drawing.Color.White
        Me.DataGridTableStyle1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle1.ForeColor = System.Drawing.Color.DimGray
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2, Me.DataGridTextBoxColumn4, Me.DataGridTextBoxColumn3})
        Me.DataGridTableStyle1.GridLineColor = System.Drawing.Color.WhiteSmoke
        Me.DataGridTableStyle1.HeaderBackColor = System.Drawing.Color.Gainsboro
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.Color.Black
        Me.DataGridTableStyle1.LinkColor = System.Drawing.Color.LightSteelBlue
        Me.DataGridTableStyle1.MappingName = "DefaultStyle"
        Me.DataGridTableStyle1.PreferredColumnWidth = 100
        Me.DataGridTableStyle1.ReadOnly = True
        Me.DataGridTableStyle1.RowHeadersVisible = False
        Me.DataGridTableStyle1.SelectionBackColor = System.Drawing.Color.LightSteelBlue
        Me.DataGridTableStyle1.SelectionForeColor = System.Drawing.Color.Black
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Project Name"
        Me.DataGridTextBoxColumn1.MappingName = "Project Name"
        Me.DataGridTextBoxColumn1.NullText = ""
        Me.DataGridTextBoxColumn1.ReadOnly = True
        Me.DataGridTextBoxColumn1.Width = 250
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "Conception Date"
        Me.DataGridTextBoxColumn2.MappingName = "Conception Date"
        Me.DataGridTextBoxColumn2.NullText = ""
        Me.DataGridTextBoxColumn2.ReadOnly = True
        Me.DataGridTextBoxColumn2.Width = 110
        '
        'DataGridTextBoxColumn4
        '
        Me.DataGridTextBoxColumn4.Format = ""
        Me.DataGridTextBoxColumn4.FormatInfo = Nothing
        Me.DataGridTextBoxColumn4.HeaderText = "Latest Build"
        Me.DataGridTextBoxColumn4.MappingName = "Latest Build"
        Me.DataGridTextBoxColumn4.NullText = ""
        Me.DataGridTextBoxColumn4.ReadOnly = True
        Me.DataGridTextBoxColumn4.Width = 110
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "Working Directory"
        Me.DataGridTextBoxColumn3.MappingName = "Working Directory"
        Me.DataGridTextBoxColumn3.NullText = ""
        Me.DataGridTextBoxColumn3.ReadOnly = True
        Me.DataGridTextBoxColumn3.Width = 268
        '
        'Label1
        '
        Me.Label1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(-1, -1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(502, 31)
        Me.Label1.TabIndex = 24
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label1, "Project description")
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.RichTextBox1)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(8, 232)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(500, 304)
        Me.Panel2.TabIndex = 25
        Me.ToolTip1.SetToolTip(Me.Panel2, "Project description")
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.White
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.DetectUrls = False
        Me.RichTextBox1.ForeColor = System.Drawing.Color.Black
        Me.RichTextBox1.Location = New System.Drawing.Point(8, 72)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(480, 224)
        Me.RichTextBox1.TabIndex = 25
        Me.RichTextBox1.Text = ""
        Me.ToolTip1.SetToolTip(Me.RichTextBox1, "Project description")
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.DarkSlateBlue
        Me.Label3.Location = New System.Drawing.Point(8, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 14)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Browse Project Folder"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.Label3, "Open the active project window")
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.DarkSlateBlue
        Me.Label2.Location = New System.Drawing.Point(8, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 14)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Install Application"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.Label2, "Launch the application installer if located")
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(518, 232)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(250, 250)
        Me.PictureBox1.TabIndex = 26
        Me.PictureBox1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox1, "Project preview image")
        '
        'buildlabel
        '
        Me.buildlabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.buildlabel.BackColor = System.Drawing.Color.White
        Me.buildlabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.buildlabel.ForeColor = System.Drawing.Color.Black
        Me.buildlabel.Location = New System.Drawing.Point(8, 544)
        Me.buildlabel.Name = "buildlabel"
        Me.buildlabel.Size = New System.Drawing.Size(500, 23)
        Me.buildlabel.TabIndex = 28
        Me.buildlabel.Text = " Build Number: Not Available"
        Me.buildlabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.buildlabel, "Application Version Number")
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Location = New System.Drawing.Point(184, 14)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(49, 49)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 2
        Me.PictureBox2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox2, "Application Icon")
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.DataGrid1)
        Me.Panel1.Location = New System.Drawing.Point(8, 56)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(760, 168)
        Me.Panel1.TabIndex = 23
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.PictureBox2)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Location = New System.Drawing.Point(518, 490)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(250, 77)
        Me.Panel3.TabIndex = 26
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(432, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "PROJECT FOLDER"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the project directory below"
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(8, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(416, 32)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "USE THE DATA GRID BELOW TO BROWSE THROUGH THE VARIOUS PROJECTS CONTAINED IN THE ‘" & _
        "PROJECT FOLDER’ DIRECTORY. ONCE YOU HAVE SELECTED A PROJECT IN THE GRID, ITS DET" & _
        "AILS WILL APPEAR BELOW AS WELL AS A LINK TO ITS INSTALLER AND PROJECT DIRECTORY." & _
        ""
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'DataSet1
        '
        Me.DataSet1.DataSetName = "NewDataSet"
        Me.DataSet1.Locale = New System.Globalization.CultureInfo("en-ZA")
        Me.DataSet1.Tables.AddRange(New System.Data.DataTable() {Me.DataTable1})
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(778, 576)
        Me.Controls.Add(Me.buildlabel)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.FullError)
        Me.Controls.Add(Me.ButtonFolderBrowse)
        Me.Controls.Add(Me.txtBaseFolder)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel3)
        Me.ForeColor = System.Drawing.Color.Gray
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(784, 608)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Project Menu          (Build 20060215.5)"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dv, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataTable1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try

            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()

        Catch ex As Exception
            MsgBox("An error occurred in Work Project Menu's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal exception As Exception)
        Try
            Dim message As String
            If FullError.Checked = True Then
                message = exception.ToString
            Else
                message = exception.Message.ToString
            End If

            If Not (exception.GetType.Name = "ThreadAbortException") Then
                Dim Display_Message1 As New Display_Message(message)
                Display_Message1.ShowDialog()
            End If

        Catch ex As Exception
            MsgBox("An error occurred in Work Project Menu's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub ButtonFolderBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFolderBrowse.Click
        Try
            Dim result As DialogResult
            If Not txtBaseFolder.Text = "" And txtBaseFolder.Text Is Nothing = False Then

                Dim filecheck As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(txtBaseFolder.Text)
                If filecheck.Exists = True Then
                    FolderBrowserDialog1.SelectedPath = txtBaseFolder.Text
                End If
            End If
            result = FolderBrowserDialog1.ShowDialog()
            If result = DialogResult.OK Or result = DialogResult.Yes Then
                txtBaseFolder.Text = FolderBrowserDialog1.SelectedPath
                load_data()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Public Sub WorkerProjectFoundHandler(ByVal filename As String, ByVal queue As String)
        Try
            Dim t As Label = New Label
            t.Height = 8
            t.Text = filename
            Panel1.Controls.Add(t)
            Panel1.Refresh()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Save_Registry_Values()
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub

   

    Private Sub load_project_data(ByVal myrow As Integer, ByVal mycolumn As Integer)
        Try

        
        clear_project_data()
        DataGrid1.CurrentCell = New DataGridCell(myrow, mycolumn)
        DataGrid1.Select(myrow)
        Label1.Text = "  " & DataGrid1.Item(myrow, 0).ToString()
            Dim finfo As FileInfo = New FileInfo(DataGrid1.Item(myrow, 3).ToString() & "\Source Code\" & DataGrid1.Item(myrow, 0).ToString() & " Installer\Images\installer_background.jpg")
        '   MsgBox(finfo.FullName)
        If finfo.Exists = True Then
            '      MsgBox("accepted")
            Dim i As Image
            i = i.FromFile(finfo.FullName)
            '     MsgBox("image created")
            Label1.BackgroundImage = i
            Label1.Width = i.Width
            Label1.Height = i.Height

        Else
            Label1.Height = 63
            Label1.Width = 500
            Label1.BackgroundImage = Nothing
            Label1.BackColor = Color.White
        End If

        Label1.Refresh()

            finfo = New FileInfo(DataGrid1.Item(myrow, 3).ToString() & "\preview_image.jpg")
        If finfo.Exists = True Then
                PictureBox1.Image = Image.FromFile(DataGrid1.Item(myrow, 3).ToString() & "\preview_image.jpg")
            PictureBox1.Update()
        End If

            finfo = New FileInfo(DataGrid1.Item(myrow, 3).ToString() & "\Source Code\" & DataGrid1.Item(myrow, 0).ToString() & " Installer\Images\Application_Icon.jpg")
            If finfo.Exists = True Then
                PictureBox2.Image = Image.FromFile(finfo.FullName)
                PictureBox2.Update()
            End If

            finfo = New FileInfo(DataGrid1.Item(myrow, 3).ToString() & "\description.txt")
            If finfo.Exists = True Then
                RichTextBox1.LoadFile(finfo.FullName, RichTextBoxStreamType.PlainText)
            End If

            finfo = New FileInfo(DataGrid1.Item(myrow, 3).ToString() & "\build.txt")
            If finfo.Exists = True Then
                Dim filereader As StreamReader = New StreamReader(finfo.FullName)
                If filereader.Peek > -1 Then
                    buildlabel.Text = "  Build Number: " & filereader.ReadLine
                Else
                    buildlabel.Text = "  Build Number: Not Available"
                End If
                filereader.Close()
                filereader = Nothing
            Else
                buildlabel.Text = "  Build Number: Not Available"
            End If

            Dim dinfo As DirectoryInfo
            dinfo = New DirectoryInfo(DataGrid1.Item(myrow, 3).ToString() & "\")
            Dim dinfo2 As DirectoryInfo
            dinfo2 = New DirectoryInfo(DataGrid1.Item(myrow, 3).ToString() & "\Release\")
            If dinfo.Exists = True Then
                Label3.Enabled = True
                runprog2 = "explorer """ & DataGrid1.Item(myrow, 3).ToString() & "\" & """"

                finfo = New FileInfo(DataGrid1.Item(myrow, 3).ToString() & "\Release\" & DataGrid1.Item(myrow, 0).ToString() & " Installer.exe")
                Label2.Enabled = True
                If finfo.Exists = True Then
                    runprog1 = """" & finfo.FullName & """"
                Else
                    If dinfo2.Exists = True Then
                        runprog1 = "explorer """ & DataGrid1.Item(myrow, 3).ToString() & "\Release\" & """"
                    Else
                        runprog1 = runprog2
                    End If
                End If
            End If
            finfo = Nothing
            dinfo2 = Nothing
            dinfo = Nothing

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    'Private Sub dataGrid1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGrid1.MouseUp
    '    Try
    '        Dim pt = New Point(e.X, e.Y)
    '        Dim hti As DataGrid.HitTestInfo = DataGrid1.HitTest(pt)
    '        If hti.Type = DataGrid.HitTestType.Cell Then
    '            load_project_data(hti.Row, hti.Column)
    '        End If
    '    Catch ex As Exception
    '        Error_Handler(ex)
    '    End Try
    'End Sub


    Private Sub dataGrid1_KeyDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged
        Try
            load_project_data(DataGrid1.CurrentCell.RowNumber, DataGrid1.CurrentCell.ColumnNumber)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Load_Registry_Values()
            load_data()
            DataGridTableStyle1.MappingName = DataTable1.TableName
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub load_data()
        Try
            DataSet1.Tables(0).Clear()
            clear_project_data()

            Dim dinfo As DirectoryInfo = New DirectoryInfo(txtBaseFolder.Text)
            If dinfo.Exists = True Then


                Dim projectcount As Integer
                Dim direct As DirectoryInfo
                projectcount = 0

                For Each direct In dinfo.GetDirectories()
                    Dim dr As DataRow = DataSet1.Tables(0).NewRow()
                    Dim projectname As String
                    projectname = ""
                    Dim tempprojectname() As String
                    tempprojectname = direct.Name.Split(" ")
                    If tempprojectname.Length > 3 Then
                        Dim i, dashpos As Integer
                        For i = 0 To (tempprojectname.Length - 1)
                            If tempprojectname(i) = "-" Then
                                If i > 0 Then
                                    dashpos = i - 1
                                End If
                            End If
                        Next
                        For i = 0 To dashpos
                            projectname = projectname + " " + tempprojectname(i)
                        Next
                        If tempprojectname(tempprojectname.Length - 4) = "-" Then

                            Dim dte As Date
                            dte = dte.Parse(tempprojectname(tempprojectname.Length - 3) & " " & tempprojectname(tempprojectname.Length - 2) & " " & tempprojectname(tempprojectname.Length - 1))
                            dr.Item(0) = projectname.Trim()
                            dr.Item(1) = Format(dte, "yyyy.MM.dd")
                            dr.Item(3) = direct.FullName




                            Dim binfo As FileInfo = New FileInfo((direct.FullName & "\build.txt").Replace("\\", "\"))
                            If binfo.Exists = True Then
                                Dim filereader As StreamReader = New StreamReader(binfo.FullName)
                                If filereader.Peek > -1 Then
                                    dr.Item(2) = filereader.ReadLine
                                Else
                                    dr.Item(2) = "00000000.0"
                                End If
                                filereader.Close()
                                filereader = Nothing
                            Else
                                dr.Item(2) = "00000000.0"
                            End If
                            binfo = Nothing



                            DataSet1.Tables(0).Rows.Add(dr)
                        End If
                        If tempprojectname(tempprojectname.Length - 3) = "-" Then

                            Dim dte As Date
                            dte = dte.Parse("01 " & tempprojectname(tempprojectname.Length - 2) & " " & tempprojectname(tempprojectname.Length - 1))
                            dr.Item(0) = projectname.Trim()
                            dr.Item(1) = Format(dte, "yyyy.MM.dd")
                            dr.Item(3) = direct.FullName
                            Dim binfo As FileInfo = New FileInfo((direct.FullName & "\build.txt").Replace("\\", "\"))
                            If binfo.Exists = True Then
                                Dim filereader As StreamReader = New StreamReader(binfo.FullName)
                                If filereader.Peek > -1 Then
                                    dr.Item(2) = filereader.ReadLine
                                Else
                                    dr.Item(2) = "00000000.0"
                                End If
                                filereader.Close()
                                filereader = Nothing
                            Else
                                dr.Item(2) = "00000000.0"
                            End If
                            binfo = Nothing

                            DataSet1.Tables(0).Rows.Add(dr)
                        End If


                    End If
                Next
                direct = Nothing
            End If
            dinfo = Nothing
            If DataSet1.Tables(0).Rows.Count = 1 Then
                DataGrid1.CaptionText = "Work Projects   (" & DataSet1.Tables(0).Rows.Count & " project loaded)"
            Else
                DataGrid1.CaptionText = "Work Projects   (" & DataSet1.Tables(0).Rows.Count & " projects loaded)"
            End If

            If DataSet1.Tables(0).Rows.Count > 0 Then
                load_project_data(0, 0)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub clear_project_data()
        Try
            runprog1 = ""
            runprog2 = ""
            Label1.Text = ""
            Label1.BackgroundImage = Nothing
            Label1.Refresh()
            RichTextBox1.Clear()
            RichTextBox1.Refresh()
            Label2.Enabled = False
            Label3.Enabled = False
            PictureBox1.Image = Nothing
            PictureBox1.Update()
            PictureBox2.Image = Nothing
            PictureBox2.Update()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Work Project Menu") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Work Project Menu")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Work Project Menu", True)

            str = oKey.GetValue("projectdirectory")
            If Not IsNothing(str) And Not (str = "") Then
                projectdirectory = str
            Else
                configflag = True
                oKey.SetValue("projectdirectory", (Application.StartupPath & "\").Replace("\\", "\"))
                projectdirectory = (Application.StartupPath & "\").Replace("\\", "\")
            End If


            oKey.Close()
            oReg.Close()

            txtBaseFolder.Text = projectdirectory

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Work Project Menu", True)

            oKey.SetValue("projectdirectory", txtBaseFolder.Text)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            myProcess.Dispose()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching DOS shell. The program will attempt to recover shortly.")
        End Try
        Return s
    End Function

    Private Sub Label3_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.MouseHover
        Try
            Label3.ForeColor = Color.Firebrick
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Label3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.MouseLeave
        Try
            Label3.ForeColor = Color.DarkSlateBlue
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Label2_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.MouseHover
        Try
            Label2.ForeColor = Color.Firebrick
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Label2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        Try
            Label2.ForeColor = Color.DarkSlateBlue
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Try
            If Not runprog1 = "" Then
                DosShellCommand(runprog1)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Try
            If Not runprog2 = "" Then
                DosShellCommand(runprog2)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


End Class
