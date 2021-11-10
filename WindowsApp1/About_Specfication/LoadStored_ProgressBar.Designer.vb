<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class LoadStored_ProgressBar_Form
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LoadStored_ProgressBar_Form))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Done_Button = New System.Windows.Forms.Button()
        Me.LoadStored_Timer = New System.Windows.Forms.Timer(Me.components)
        Me.SQLite_TotalDataLoading_Label = New System.Windows.Forms.Label()
        Me.SQLite_LoadingText_Label = New System.Windows.Forms.Label()
        Me.SQLite_EachDataLoading_Label = New System.Windows.Forms.Label()
        Me.SQLite_Loading_PictureBox = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SQLite_Loading_PictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.Image = Global.WindowsApp1.My.Resources.Resources.safe_image
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(97, 77)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(144, 83)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Done_Button
        '
        Me.Done_Button.BackgroundImage = CType(resources.GetObject("Done_Button.BackgroundImage"), System.Drawing.Image)
        Me.Done_Button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Done_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Done_Button.Location = New System.Drawing.Point(132, 166)
        Me.Done_Button.Name = "Done_Button"
        Me.Done_Button.Size = New System.Drawing.Size(75, 31)
        Me.Done_Button.TabIndex = 1
        Me.Done_Button.UseVisualStyleBackColor = True
        '
        'SQLite_TotalDataLoading_Label
        '
        Me.SQLite_TotalDataLoading_Label.AutoSize = True
        Me.SQLite_TotalDataLoading_Label.Font = New System.Drawing.Font("微軟正黑體", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.SQLite_TotalDataLoading_Label.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.SQLite_TotalDataLoading_Label.Location = New System.Drawing.Point(218, 8)
        Me.SQLite_TotalDataLoading_Label.Name = "SQLite_TotalDataLoading_Label"
        Me.SQLite_TotalDataLoading_Label.Size = New System.Drawing.Size(95, 40)
        Me.SQLite_TotalDataLoading_Label.TabIndex = 71
        Me.SQLite_TotalDataLoading_Label.Text = "/ 111"
        '
        'SQLite_LoadingText_Label
        '
        Me.SQLite_LoadingText_Label.AutoSize = True
        Me.SQLite_LoadingText_Label.Font = New System.Drawing.Font("微軟正黑體", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.SQLite_LoadingText_Label.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.SQLite_LoadingText_Label.Location = New System.Drawing.Point(44, 16)
        Me.SQLite_LoadingText_Label.Name = "SQLite_LoadingText_Label"
        Me.SQLite_LoadingText_Label.Size = New System.Drawing.Size(87, 24)
        Me.SQLite_LoadingText_Label.TabIndex = 73
        Me.SQLite_LoadingText_Label.Text = "Loading:"
        '
        'SQLite_EachDataLoading_Label
        '
        Me.SQLite_EachDataLoading_Label.Font = New System.Drawing.Font("微軟正黑體", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.SQLite_EachDataLoading_Label.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.SQLite_EachDataLoading_Label.Location = New System.Drawing.Point(129, 8)
        Me.SQLite_EachDataLoading_Label.Name = "SQLite_EachDataLoading_Label"
        Me.SQLite_EachDataLoading_Label.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SQLite_EachDataLoading_Label.Size = New System.Drawing.Size(86, 40)
        Me.SQLite_EachDataLoading_Label.TabIndex = 72
        Me.SQLite_EachDataLoading_Label.Text = "111"
        '
        'SQLite_Loading_PictureBox
        '
        Me.SQLite_Loading_PictureBox.Image = CType(resources.GetObject("SQLite_Loading_PictureBox.Image"), System.Drawing.Image)
        Me.SQLite_Loading_PictureBox.Location = New System.Drawing.Point(3, 3)
        Me.SQLite_Loading_PictureBox.Name = "SQLite_Loading_PictureBox"
        Me.SQLite_Loading_PictureBox.Size = New System.Drawing.Size(50, 50)
        Me.SQLite_Loading_PictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SQLite_Loading_PictureBox.TabIndex = 70
        Me.SQLite_Loading_PictureBox.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SQLite_LoadingText_Label)
        Me.Panel1.Controls.Add(Me.SQLite_Loading_PictureBox)
        Me.Panel1.Controls.Add(Me.SQLite_TotalDataLoading_Label)
        Me.Panel1.Controls.Add(Me.SQLite_EachDataLoading_Label)
        Me.Panel1.Location = New System.Drawing.Point(8, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(322, 59)
        Me.Panel1.TabIndex = 74
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(39, 203)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(202, 70)
        Me.TextBox1.TabIndex = 75
        '
        'LoadStored_ProgressBar_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(342, 346)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Done_Button)
        Me.Controls.Add(Me.PictureBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "LoadStored_ProgressBar_Form"
        Me.Text = "LoadStored_ProgressBar"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SQLite_Loading_PictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Done_Button As Button
    Friend WithEvents LoadStored_Timer As Timer
    Friend WithEvents SQLite_TotalDataLoading_Label As Label
    Friend WithEvents SQLite_LoadingText_Label As Label
    Friend WithEvents SQLite_EachDataLoading_Label As Label
    Friend WithEvents SQLite_Loading_PictureBox As PictureBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TextBox1 As TextBox
End Class
