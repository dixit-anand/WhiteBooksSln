Option Explicit Off
'Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection.Emit
Imports System.Windows.Forms.Design.AxImporter
Imports ADODB
' username = U1-45682292738020185013354806
' pwd = 1
Public Class LOGIN_SCREEN
    Dim ORANGECONN_SERVER, ORANGECONN_LOCAL As New ADODB.Connection
    Dim USERINFO As New ADODB.Recordset
    Dim USERCOMPANIES As New ADODB.Recordset
    Dim GODOWN As New ADODB.Recordset
    Dim VERINFO As New ADODB.Recordset
    Dim DATACENTRE As String
    Dim IP_ADDRESS As String
    Dim GLB_UID As String
    Dim GLB_PWD As String
    Dim GLB_UNAME As String
    Dim PSWD1 As String
    Dim PSWD2 As String
    Dim LOGINDATALIST As New DataTable
    Private Sub OK_BUTTON_Click(sender As Object, e As EventArgs) Handles OK_BUTTON.Click
        If USERCOMPANIES.State = ObjectStateEnum.adStateOpen Then
            USERCOMPANIES.Close()
        End If
        USERCOMPANIES.Open("SELECT * FROM USERPERMISSIONS WHERE ID = 62", ORANGECONN_SERVER, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
        Dim ADAPTER = New OleDb.OleDbDataAdapter()
        ADAPTER.Fill(LOGINDATALIST, USERCOMPANIES)
        Dim UID = LOGINDATALIST.Rows.Item(0).Item("USER_ID")
        Dim CID = LOGINDATALIST.Rows.Item(0).Item("COMPANY_ID")
        MsgBox("COMP-ID = " + CID + "USER-ID = " + UID)
        MAIN_SCREEN.Show()
    End Sub

    Private Sub CANCEL_Click(sender As Object, e As EventArgs) Handles CANCEL.Click
        Application.Exit()
    End Sub

    Private Sub LOGIN_SCREEN_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim GLB_VERSION_NUMBER = 9

        Dim GDC, DBASENAME, GLB_FINANCIAL_YEAR, GLB_YEAR_BEGINNING_DATE, GLB_YEAR_END_DATE, DATABASE_NAME, GLB_CONNSTRING_CLIENTS
        On Error GoTo ERRMSG
        Drive1.DataSource = My.Computer.FileSystem.Drives

        If ORANGECONN_SERVER.State = ConnectionState.Open Then
            ORANGECONN_SERVER.Close()
        End If
        If ORANGECONN_LOCAL.State = ConnectionState.Open Then
            ORANGECONN_LOCAL.Close()
        End If
        'GDC = GETDATACENTRE()
        GETDATACENTRE()

        DBASENAME = "C:\ORANGE_BUSINESS_SOLUTION\ORANGEREPORT.accdb"

        GLB_FINANCIAL_YEAR = "24-25" ' for dd tyres
        If GLB_FINANCIAL_YEAR = "24-25" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2024"
            GLB_YEAR_END_DATE = "31/03/2025"
        End If
        ORANGECONN_LOCAL.Open("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & DBASENAME & ";Persist Security Info=False;")
        DATABASE_NAME = "VERSIONINFO"
        If ORANGECONN_SERVER.State = ConnectionState.Open Then
            ORANGECONN_SERVER.Close()
        End If
        ORANGECONN_SERVER.CursorLocation = CursorLocationEnum.adUseClient
        ORANGECONN_SERVER.Open("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & GLB_UID & "; Password=" & GLB_PWD & ";Initial Catalog=" & DATABASE_NAME & ";Data Source=" & IP_ADDRESS & "")
        If VERINFO.State = ObjectStateEnum.adStateOpen Then
            VERINFO.Close()
        End If
        VERINFO.Open("SELECT * FROM VERINFORMATION WHERE VERSIONNO = " & GLB_VERSION_NUMBER & "", ORANGECONN_SERVER, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
        If VERINFO.EOF = True Then
            ' MsgBox(VERINFO.Fields.Item("VERSIONNO").Value)
            MsgBox("An Upgraded Version Of BUSINESS SOLUTION Is Available For You. Call Orange Help Desk On +91 9972527251 For New Copy.", vbExclamation, "Important")
            Me.Close()
        Else
            'MsgBox(VERINFO.Fields.Item("VERSIONNO").Value)
            ' MsgBox(VERINFO!VERSIONNO)
            ' MsgBox(VERINFO("VERSIONNO"))
            DATABASE_NAME = "ORANGECLIENTS"
            If ORANGECONN_SERVER.State = ObjectStateEnum.adStateOpen Then
                ORANGECONN_SERVER.Close()
            End If
            ORANGECONN_SERVER.CursorLocation = CursorLocationEnum.adUseClient
            ORANGECONN_SERVER.Open("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & GLB_UID & "; Password=" & GLB_PWD & ";Initial Catalog=" & DATABASE_NAME & ";Data Source=" & IP_ADDRESS & "")
            GLB_CONNSTRING_CLIENTS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & GLB_UID & "; Password=" & GLB_PWD & ";Initial Catalog=" & DATABASE_NAME & ";Data Source=" & IP_ADDRESS & ""
        End If
        Exit Sub
ERRMSG:
        If Err.Number = -2147467259 Then
            MsgBox("Internet Connection Is Not On. Please Check Your Modem Or Contact Network Administrator For Solution.", vbCritical, "Note")

        Else
            MsgBox(Err.Description & "Click OK to Exit ", vbCritical, "Unknown Error occured")

        End If
        Application.Exit()
    End Sub



    Public Sub GETDATACENTRE()

        DATACENTRE = "L" ' LOCAL HARD DISK
        If DATACENTRE = "G" Then

        ElseIf DATACENTRE = "A" Then
            ''''''''AMAZON WEB SERVICES'''' uptil 11/02/2018 end of day
            '        IP_ADDRESS = "orange.czz5lvlmdmld.ap-south-1.rds.amazonaws.com"
            '        GLB_UID = "ORANGEADMIN"
            '        GLB_PWD = "RSP170967RSP*"
            ' BELOW IS NEW INSTANCE CREATED ON 18/02/2018 and data transferred from orange to orangeitsolns
            ' on 22/02/2018 and new exe given on 22/02/2018 to all customers with below credentials.
            ' above credentials do not have any meaning henceforth 22/02/2018
            IP_ADDRESS = "orangeitsolns.czz5lvlmdmld.ap-south-1.rds.amazonaws.com"
            GLB_UID = "ORANGEADMIN"
            GLB_PWD = "RSP170967RSP*"
        ElseIf DATACENTRE = "L" Then
            '''''''LOCAL HARD DISK'''''
            '        IP_ADDRESS = "ORANGEITS\SQLEXPRESS" ' MS-SQL-2005 EXPECTS THIS STATEMENT
            '        IP_ADDRESS = "ORANGEITS" 'ON HP LAPTOP WITH MS-SQL-2008
            '        IP_ADDRESS = "ORANGE" 'ON COMPAQ LAPTOP WITH MS-SQL-2014
            IP_ADDRESS = "ORANGE" 'ON ASUS LAPTOP WITH MS-SQL-2014
            '        IP_ADDRESS = "DELL-PC\SQLEXPRESS" 'ON SARDARJI LAPTOPP WITH MS-SQL-2012
            '        IP_ADDRESS = "192.168.1.4" ' for orange wifi at home
            '        IP_ADDRESS = "192.168.1.89" ' for gopinath home
            '        IP_ADDRESS = "192.168.1.14" ' for shrihari
            '        IP_ADDRESS = "192.168.225.58" ' for reliance jio
            '        IP_ADDRESS = "192.168.1.104" ' kamble tyres server
            '        IP_ADDRESS = "192.168.91.162" ' for rajendra mobile
            '        IP_ADDRESS = "MRF\MSSQLSERVER"
            '        IP_ADDRESS = "localhost" ' FOR KAMBLE TYRES
            '        IP_ADDRESS = "DDTYRES\SQLEXPRESS" ' FOR DD TYRES FOR SQL SERVER 2014
            'IP_ADDRESS = "192.168.1.89" ' gopinath home.
            GLB_UID = "sa"
            GLB_PWD = "RSP170967RSP*"

        ElseIf DATACENTRE = "R" Then
            '''''''RAGIGUDDA TEMPLE'''''
            IP_ADDRESS = "GENERAL\SQLEXPRESS" ' FOR RAGIGUDDA
            GLB_UID = "sa"
            GLB_PWD = "RSP170967RSP*"
        End If
    End Sub

    'Private Sub LOGIN_SCREEN_Activated(sender As Object, e As EventArgs) Handles Me.Activated
    '    '    xx = "ABcdef"
    '    '    MsgBox UCase(xx)
    '    Me.Left = 5000
    '    Me.Top = 4000
    '    Label1.Visible = False
    '    Label2.Visible = False
    '    Label3.Visible = False
    '    Label4.Visible = False
    '    'Label7.Visible = False
    '    USER_CODE.Visible = False
    '    USER_NAME.Visible = False
    '    USER_PASSWORD.Visible = False
    '    COMPANY_COMBO.Visible = False
    '    BRANCH_COMBO.Visible = False
    '    OK_BUTTON.Enabled = True
    '    'RadioButton1.Select()
    '    'Drive1.Focus()
    'End Sub

    Private Sub LOGIN_SCREEN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If Val(e.KeyChar) = 27 Then
            Application.Exit()
        End If
    End Sub
    Private Sub USER_NAME_KeyDown(sender As Object, e As KeyEventArgs) Handles USER_NAME.KeyDown
        If e.KeyCode = Keys.Enter Then
            USER_PASSWORD.Focus()
        End If

    End Sub


    Private Sub USER_PASSWORD_KeyDown(sender As Object, e As KeyEventArgs) Handles USER_PASSWORD.KeyDown
        If e.KeyCode = Keys.Enter Then
            If USERCOMPANIES.State = ObjectStateEnum.adStateOpen Then
                USERCOMPANIES.Close()
            End If
            USERCOMPANIES.Open("SELECT * FROM USERPERMISSIONS,COMPANY WHERE COMPANY.COMPANY_ID=USERPERMISSIONS.COMPANY_ID AND USER_ID = '" & USER_CODE.Text & "' AND USER_STATUS=1", ORANGECONN_SERVER, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
            If USERCOMPANIES.BOF = True And USERCOMPANIES.EOF = True Then
                MsgBox("Your Name Is Not Authorised By Any Company. Ask Company To Register And Grant You Permission To Operate The Software.", vbInformation, "Important Message")
                Me.Close()
                Exit Sub
            Else
                If USER_PASSWORD.Text = PSWD1 Then
                    GLB_PWD_SELECT = 1
                    GLB_USER_PSWD = PSWD1
                ElseIf USER_PASSWORD.Text = PSWD2 Then
                    GLB_PWD_SELECT = 2
                    GLB_USER_PSWD = PSWD2
                ElseIf USER_PASSWORD.Text <> GLB_USER_PSWD Then
                    MsgBox("The Password You Entered Is Wrong. Re-Enter Password.", vbCritical, "Note")
                    USER_PASSWORD.Text = ""
                    USER_PASSWORD.Focus()
                    Exit Sub
                End If
                COMPANY_COMBO.Items.Clear()
                Do While Not USERCOMPANIES.EOF
                    COMPANY_COMBO.Items.Add(USERCOMPANIES.Fields("company_name").Value)
                    USERCOMPANIES.MoveNext()
                Loop
                COMPANY_COMBO.SelectedIndex() = 0
                COMPANY_COMBO.Focus()

            End If
        End If
    End Sub

    Private Sub COMPANY_COMBO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles COMPANY_COMBO.SelectedIndexChanged

    End Sub

    Private Sub COMPANY_COMBO_KeyDown(sender As Object, e As KeyEventArgs) Handles COMPANY_COMBO.KeyDown
        If e.KeyCode = Keys.Enter Then
            BRANCH_COMBO.Focus()
        End If
    End Sub
    Private Sub BRANCH_COMBO_KeyDown(sender As Object, e As KeyEventArgs) Handles BRANCH_COMBO.KeyDown
        If e.KeyCode = Keys.Enter Then
            FINANCIAL_YEAR.Focus()
        End If
    End Sub
    Private Sub FINANCIAL_YEAR_KeyDown(sender As Object, e As KeyEventArgs) Handles FINANCIAL_YEAR.KeyDown
        If e.KeyCode = Keys.Enter Then
            OK_BUTTON.Focus()
        End If
    End Sub
    Private Sub USER_CODE_KeyDown(sender As Object, e As KeyEventArgs) Handles USER_CODE.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Mid(USER_CODE.Text, 1, 1) = "" Then
                USER_CODE.Focus()
                Exit Sub
            End If
            If USERINFO.State = ObjectStateEnum.adStateOpen Then
                USERINFO.Close()
            End If
            If Strings.UCase(Mid(USER_CODE.Text, 1, 1)) = "U" Then
                USERINFO.Open("SELECT * FROM USERS WHERE USER_ID = '" & Trim(USER_CODE.Text) & "'", ORANGECONN_SERVER, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                If (USERINFO.BOF = True And USERINFO.EOF = True) Then
                    MsgBox("The User Code You Entered Is Wrong. Please Try Again.", vbCritical, "Message")
                    USER_CODE.Text = ""
                    USER_CODE.Focus()
                    Exit Sub
                Else
                    If Trim(USERINFO.Fields.Item("USER_NAME").Value) = "NIL" Then  '  Or String.IsNullOrEmpty(Trim(USERNAME))
                        MsgBox("Your User Registration Details Are Not Yet Entered. Please Enter The Details In The Screen Shown Next.", vbInformation, "Message")
                        USER_REGFORM.Show()
                        Exit Sub
                    Else
                        GLB_UNAME = Trim(USERINFO.Fields.Item("USER_NAME").Value)
                        PSWD1 = Trim(USERINFO.Fields.Item("USER_PASSWORD").Value)
                        PSWD2 = Trim(USERINFO.Fields.Item("USER_PASSWORD1").Value)
                        USER_NAME.Text = GLB_UNAME
                        USER_PASSWORD.Focus()
                    End If
                End If
            ElseIf Mid(USER_CODE.Text, 1, 1) = "C" Then
                USERINFO.Open("SELECT * FROM COMPANY WHERE COMPANY_ID = '" & Trim(USER_CODE.Text) & "'", ORANGECONN_SERVER, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                If (USERINFO.BOF = True And USERINFO.EOF = True) Or Mid(USER_CODE.Text, 1, 1) <> "C" Then
                    MsgBox("The Company Code You Entered Is Wrong. Please Try Again.", vbCritical, "Message")
                    USER_CODE.Text = ""
                    USER_CODE.Focus()
                    Exit Sub
                Else
                    If Trim(USERINFO.Fields.Item("COMPANY_NAME").Value) = "NIL" Then  '  Or String.IsNullOrEmpty(Trim(USERNAME))
                        MsgBox("Your Company Registration Details Are Not Yet Entered. Please Enter The Details In The Screen Shown Next.", vbInformation, "Message")
                        COMPANY_REGFORM.Show()
                        Exit Sub
                    Else
                        MsgBox("Your Company Is Already Registered On Server. Please Click On OK To Exit.", vbInformation, "Message")
                        Application.Exit()
                        Exit Sub
                    End If
                End If
            ElseIf Mid(USER_CODE.Text, 1, 1) <> "U" Or Mid(USER_CODE.Text, 1, 1) <> "U" Then
                MsgBox("The User Code Or The Company Code You Entered Is Wrong. Please Try Again.", vbCritical, "Message")
                USER_CODE.Text = ""
                USER_CODE.Focus()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub USER_PASSWORD_TextChanged(sender As Object, e As EventArgs) Handles USER_PASSWORD.TextChanged

    End Sub
End Class