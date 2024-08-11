Option Explicit Off
Imports System.Reflection.Emit
Imports System.Windows.Forms.Design.AxImporter
Imports ADODB
Public Class LOGIN_SCREEN
    Dim USERINFO As New ADODB.Recordset
    Dim USERCOMPANIES As New ADODB.Recordset
    Dim GODOWN As New ADODB.Recordset
    Dim VERINFO As New ADODB.Recordset
    Dim DATACENTRE As String
    Dim IP_ADDRESS As String
    Dim GLB_UID As String
    Dim GLB_PWD As String
    Dim PSWD1 As String
    Dim PSWD2 As String
    Private Sub OK_BUTTON_Click(sender As Object, e As EventArgs) Handles OK_BUTTON.Click
        MAIN_SCREEN.Show()
    End Sub

    Private Sub CANCEL_Click(sender As Object, e As EventArgs) Handles CANCEL.Click
        Application.Exit()
    End Sub

    Private Sub LOGIN_SCREEN_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim GLB_VERSION_NUMBER = 9
        Dim ORANGECONN_SERVER, ORANGECONN_LOCAL As New ADODB.Connection
        Dim GDC, DBASENAME, GLB_FINANCIAL_YEAR, GLB_YEAR_BEGINNING_DATE, GLB_YEAR_END_DATE, DATABASE_NAME, GLB_CONNSTRING_CLIENTS
        On Error GoTo ERRMSG
        Drive1.DataSource = My.Computer.FileSystem.Drives

        If ORANGECONN_SERVER.State = ConnectionState.Open Then
            ORANGECONN_SERVER.Close()
        End If
        If ORANGECONN_LOCAL.State = ConnectionState.Open Then
            ORANGECONN_LOCAL.ClosOpe()
        End If
        'GDC = GETDATACENTRE()
        GETDATACENTRE()

        DBASENAME = "C:\ORANGE_BUSINESS_SOLUTION\ORANGEREPORT.accdb"

        GLB_FINANCIAL_YEAR = "21-22" ' for dd tyres
        If GLB_FINANCIAL_YEAR = "15-16" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2015"
            GLB_YEAR_END_DATE = "31/03/2016"
        ElseIf GLB_FINANCIAL_YEAR = "16-17" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2016"
            GLB_YEAR_END_DATE = "31/03/2017"
        ElseIf GLB_FINANCIAL_YEAR = "17-18" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2017"
            GLB_YEAR_END_DATE = "31/03/2018"
        ElseIf GLB_FINANCIAL_YEAR = "18-19" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2018"
            GLB_YEAR_END_DATE = "31/03/2019"
        ElseIf GLB_FINANCIAL_YEAR = "19-20" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2019"
            GLB_YEAR_END_DATE = "31/03/2020"
        ElseIf GLB_FINANCIAL_YEAR = "20-21" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2020"
            GLB_YEAR_END_DATE = "31/03/2021"
        ElseIf GLB_FINANCIAL_YEAR = "21-22" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2021"
            GLB_YEAR_END_DATE = "31/03/2022"
        ElseIf GLB_FINANCIAL_YEAR = "22-23" Then
            GLB_YEAR_BEGINNING_DATE = "01/04/2022"
            GLB_YEAR_END_DATE = "31/03/2023"

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
            MsgBox("An Upgraded Version Of BUSINESS SOLUTION Is Available For You. Call Orange Help Desk On +91 9972527251 For New Copy.", vbExclamation, "Important")
            Me.Close()
        Else
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
        '    DATACENTRE = "N" ' NXTGEN
        '    DATACENTRE = "G" ' GODADDY
        '     DATACENTRE = "A" ' AMAZON
        DATACENTRE = "L" ' LOCAL HARD DISK
        '    DATACENTRE = "R" ' RAGIGUDDA TEMPLE

        If DATACENTRE = "G" Then
            ''''''''''''''GODADDY SERVER''''''''
            IP_ADDRESS = "203.124.106.175\SQLEXPRESS" ' GODADDY server
            GLB_UID = "sa"
            GLB_PWD = "RSP170967RSP*"
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

        ElseIf DATACENTRE = "N" Then
            '''''''NXTGEN''''''''''''
            IP_ADDRESS = "45.118.181142."
            GLB_UID = "sa"
            GLB_PWD = "ORP@ssw0rd123"

        ElseIf DATACENTRE = "L" Then
            '''''''LOCAL HARD DISK'''''
            '        IP_ADDRESS = "ORANGEITS\SQLEXPRESS" ' MS-SQL-2005 EXPECTS THIS STATEMENT
            '        IP_ADDRESS = "ORANGEITS" 'ON HP LAPTOP WITH MS-SQL-2008
            '        IP_ADDRESS = "ORANGE" 'ON COMPAQ LAPTOP WITH MS-SQL-2014
            IP_ADDRESS = "ORANGE1" 'ON ASUS LAPTOP WITH MS-SQL-2014
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

    Private Sub LOGIN_SCREEN_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        '    xx = "ABcdef"
        '    MsgBox UCase(xx)
        Me.Left = 5000
        Me.Top = 4000
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        'Label7.Visible = False
        USER_CODE.Visible = False
        USER_NAME.Visible = False
        USER_PASSWORD.Visible = False
        COMPANY_COMBO.Visible = False
        BRANCH_COMBO.Visible = False
        OK_BUTTON.Enabled = True
        RadioButton1.Select()
        Drive1.Focus()
    End Sub
End Class