<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LOGIN_SCREEN
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        BRANCH_COMBO = New ComboBox()
        USER_NAME = New TextBox()
        RadioButton1 = New RadioButton()
        Label1 = New Label()
        Label2 = New Label()
        RadioButton2 = New RadioButton()
        COMPANY_COMBO = New ComboBox()
        Label3 = New Label()
        Label4 = New Label()
        Label5 = New Label()
        USER_PASSWORD = New TextBox()
        Label6 = New Label()
        FINANCIAL_YEAR = New ComboBox()
        OK_BUTTON = New Button()
        CANCEL = New Button()
        USER_CODE = New TextBox()
        Drive1 = New ListBox()
        SuspendLayout()
        ' 
        ' BRANCH_COMBO
        ' 
        BRANCH_COMBO.DropDownStyle = ComboBoxStyle.DropDownList
        BRANCH_COMBO.FormattingEnabled = True
        BRANCH_COMBO.Location = New Point(187, 141)
        BRANCH_COMBO.Margin = New Padding(3, 2, 3, 2)
        BRANCH_COMBO.Name = "BRANCH_COMBO"
        BRANCH_COMBO.Size = New Size(367, 23)
        BRANCH_COMBO.TabIndex = 0
        ' 
        ' USER_NAME
        ' 
        USER_NAME.Location = New Point(187, 59)
        USER_NAME.Margin = New Padding(3, 2, 3, 2)
        USER_NAME.Name = "USER_NAME"
        USER_NAME.Size = New Size(367, 23)
        USER_NAME.TabIndex = 1
        ' 
        ' RadioButton1
        ' 
        RadioButton1.AutoSize = True
        RadioButton1.Checked = True
        RadioButton1.Location = New Point(18, 9)
        RadioButton1.Margin = New Padding(3, 2, 3, 2)
        RadioButton1.Name = "RadioButton1"
        RadioButton1.Size = New Size(107, 19)
        RadioButton1.TabIndex = 2
        RadioButton1.TabStop = True
        RadioButton1.Text = "Login Manually"
        RadioButton1.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(18, 141)
        Label1.Name = "Label1"
        Label1.Size = New Size(105, 15)
        Label1.TabIndex = 3
        Label1.Text = "Select Your Branch"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(18, 35)
        Label2.Name = "Label2"
        Label2.Size = New Size(124, 15)
        Label2.TabIndex = 4
        Label2.Text = "User / Company Code"
        ' 
        ' RadioButton2
        ' 
        RadioButton2.AutoSize = True
        RadioButton2.Location = New Point(187, 9)
        RadioButton2.Margin = New Padding(3, 2, 3, 2)
        RadioButton2.Name = "RadioButton2"
        RadioButton2.Size = New Size(115, 19)
        RadioButton2.TabIndex = 5
        RadioButton2.Text = "Login Thru Drive."
        RadioButton2.UseVisualStyleBackColor = True
        ' 
        ' COMPANY_COMBO
        ' 
        COMPANY_COMBO.DropDownStyle = ComboBoxStyle.DropDownList
        COMPANY_COMBO.FormattingEnabled = True
        COMPANY_COMBO.Location = New Point(187, 112)
        COMPANY_COMBO.Margin = New Padding(3, 2, 3, 2)
        COMPANY_COMBO.Name = "COMPANY_COMBO"
        COMPANY_COMBO.Size = New Size(367, 23)
        COMPANY_COMBO.TabIndex = 6
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(18, 118)
        Label3.Name = "Label3"
        Label3.Size = New Size(93, 15)
        Label3.TabIndex = 7
        Label3.Text = "Select Company"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(18, 91)
        Label4.Name = "Label4"
        Label4.Size = New Size(87, 15)
        Label4.TabIndex = 8
        Label4.Text = "Enter Password"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(18, 62)
        Label5.Name = "Label5"
        Label5.Size = New Size(65, 15)
        Label5.TabIndex = 9
        Label5.Text = "User Name"
        ' 
        ' USER_PASSWORD
        ' 
        USER_PASSWORD.Location = New Point(187, 83)
        USER_PASSWORD.Margin = New Padding(3, 2, 3, 2)
        USER_PASSWORD.Name = "USER_PASSWORD"
        USER_PASSWORD.PasswordChar = "*"c
        USER_PASSWORD.Size = New Size(367, 23)
        USER_PASSWORD.TabIndex = 10
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(18, 167)
        Label6.Name = "Label6"
        Label6.Size = New Size(79, 15)
        Label6.TabIndex = 12
        Label6.Text = "Financial Year"
        ' 
        ' FINANCIAL_YEAR
        ' 
        FINANCIAL_YEAR.DropDownStyle = ComboBoxStyle.DropDownList
        FINANCIAL_YEAR.FormattingEnabled = True
        FINANCIAL_YEAR.Location = New Point(187, 167)
        FINANCIAL_YEAR.Margin = New Padding(3, 2, 3, 2)
        FINANCIAL_YEAR.Name = "FINANCIAL_YEAR"
        FINANCIAL_YEAR.Size = New Size(367, 23)
        FINANCIAL_YEAR.TabIndex = 11
        ' 
        ' OK_BUTTON
        ' 
        OK_BUTTON.Location = New Point(187, 192)
        OK_BUTTON.Margin = New Padding(3, 2, 3, 2)
        OK_BUTTON.Name = "OK_BUTTON"
        OK_BUTTON.Size = New Size(82, 23)
        OK_BUTTON.TabIndex = 13
        OK_BUTTON.Text = "&OK"
        OK_BUTTON.UseVisualStyleBackColor = True
        ' 
        ' CANCEL
        ' 
        CANCEL.Location = New Point(472, 192)
        CANCEL.Margin = New Padding(3, 2, 3, 2)
        CANCEL.Name = "CANCEL"
        CANCEL.Size = New Size(82, 23)
        CANCEL.TabIndex = 14
        CANCEL.Text = "&Cancel"
        CANCEL.UseVisualStyleBackColor = True
        ' 
        ' USER_CODE
        ' 
        USER_CODE.Location = New Point(187, 32)
        USER_CODE.Margin = New Padding(3, 2, 3, 2)
        USER_CODE.Name = "USER_CODE"
        USER_CODE.Size = New Size(367, 23)
        USER_CODE.TabIndex = 15
        ' 
        ' Drive1
        ' 
        Drive1.FormattingEnabled = True
        Drive1.ItemHeight = 15
        Drive1.Location = New Point(187, 232)
        Drive1.Margin = New Padding(3, 2, 3, 2)
        Drive1.Name = "Drive1"
        Drive1.Size = New Size(367, 64)
        Drive1.TabIndex = 16
        ' 
        ' LOGIN_SCREEN
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(586, 314)
        ControlBox = False
        Controls.Add(Drive1)
        Controls.Add(CANCEL)
        Controls.Add(OK_BUTTON)
        Controls.Add(Label6)
        Controls.Add(FINANCIAL_YEAR)
        Controls.Add(USER_PASSWORD)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(COMPANY_COMBO)
        Controls.Add(RadioButton2)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(RadioButton1)
        Controls.Add(USER_NAME)
        Controls.Add(BRANCH_COMBO)
        Controls.Add(Label5)
        Controls.Add(USER_CODE)
        Margin = New Padding(3, 2, 3, 2)
        Name = "LOGIN_SCREEN"
        StartPosition = FormStartPosition.CenterScreen
        Text = "LOGIN SCREEN"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents BRANCH_COMBO As ComboBox
    Friend WithEvents USER_NAME As TextBox
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents COMPANY_COMBO As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents USER_PASSWORD As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents FINANCIAL_YEAR As ComboBox
    Friend WithEvents OK_BUTTON As Button
    Friend WithEvents CANCEL As Button
    Friend WithEvents USER_CODE As TextBox
    Friend WithEvents Drive1 As ListBox
End Class
