Imports System.DirectoryServices
Imports System.IO
Imports Microsoft.Win32


Public Class Form1
    Inherits System.Windows.Forms.Form
    Public Const SPI_SETDESKWALLPAPER = 20
    Const SPIF_SENDWININICHANGE = &H2
    Const SPIF_UPDATEINIFILE = &H1
    Dim ComName
    Dim ComNum

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
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(8, 8)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(736, 64)
        Me.ProgressBar1.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(8, 80)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(736, 200)
        Me.TextBox1.TabIndex = 1
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 10000
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(284, 288)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(184, 24)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Cancel"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(752, 318)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Loading...."
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Button1.Enabled = False
        Timer1.Enabled = False


        Dim mainthread As New Threading.Thread(AddressOf gogogo)
        mainthread.Start()
    End Sub
    Public Sub gogogo()
        TextBox1.Text = "Getting IP Address........"
        Dim ipaddress = getipaddress()

        If ipaddress <> "" Then
            ProgressBar1.Value += 5
            TextBox1.Text += "Got IP Address: " & ipaddress & vbNewLine
            ProgressBar1.Value += 5
            TextBox1.Text += "Determining Name..........."
            ProgressBar1.Value += 5
            Dim getnumber As System.Text.RegularExpressions.Regex
            ComNum = getnumber.Match(ipaddress, "\d*\.\d*\.\d*\.(?<fps>\d*)").Result("${fps}")
            ComName = "COM" & ComNum
            TextBox1.Text += vbNewLine & "Attemping to apply registry settings for" & ComName & "......" & vbNewLine & vbNewLine
            reg()
            TextBox1.Text += vbNewLine & "Reg cahnge complete......" & vbNewLine & vbNewLine
            ProgressBar1.Value += 5
            TextBox1.Text += "Configuring computer as " & ComName & vbNewLine
            TextBox1.Text += "Attempting to creating user:" & ComName & "..........."
            Dim namethread As New System.Threading.Thread(AddressOf changeusername)
            ProgressBar1.Value += 5
            namethread.Start()
            namethread.Join()
            ProgressBar1.Value += 5
            TextBox1.Text += "Attempting to rename computer........."
            Dim comthread As New System.Threading.Thread(AddressOf changecomname)
            ProgressBar1.Value += 5
            comthread.Start()
            comthread.Join()
            shield()
            TextBox1.Text += vbNewLine & "Setting ScreenShield PC Number............Done"
            ProgressBar1.Value += 5
            TextBox1.Text += vbNewLine & "__________________________________________________________________"
            TextBox1.ForeColor = Color.DarkBlue
            TextBox1.Text += vbNewLine & "Starting Image Changing......" & vbNewLine & vbNewLine
            TextBox1.ForeColor = Color.Black
            ProgressBar1.Value += 5
            backimg()
            TextBox1.Text += "Rebooting...................."
            ProgressBar1.Value = 100
            TextBox1.ForeColor = Color.Green
            TextBox1.Text += "Rebooting........."
            System.Threading.Thread.Sleep(10000)

            reboot()

        End If


    End Sub
    Public Sub backimg()
        TextBox1.Text += "Changing Screenshiled Image"
        File.Delete("C:\program files\screenshield\back_img.jpg")
        File.Copy("C:\program files\screenshield\images\" & ComNum & ".jpg", "C:\program files\screenshield\back_img.jpg")

        TextBox1.Text += vbNewLine & "Changing Desktop Image"
        Dim strBitmapImage As String
        strBitmapImage = "C:\WINDOWS\images\" & ComNum & ".bmp"
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop", True)
        key.SetValue("Wallpaper", strBitmapImage.ToString)
        key.SetValue("TileWallpaper", 0.ToString)
        key.SetValue("WallpaperStyle", 2.ToString)
        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

    End Sub
    Public Sub changeusername()
        Try
            Dim AD As DirectoryEntry = New DirectoryEntry("WinNT://" + Environment.MachineName + ",computer")
            Dim NewUser As DirectoryEntry = AD.Children.Find("user", "user")
            NewUser.Rename(ComName)

            NewUser.CommitChanges()
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", True)
            key.SetValue("AltDefaultUserName", ComName.ToString)
            key.SetValue("DefaultUserName", ComName.ToString)
            key.SetValue("Shell", "Explorer.exe")
            TextBox1.Text += "User account created successfully" & vbNewLine


        Catch ex As Exception
            TextBox1.Text += vbNewLine + "*********** ERROR - CAN NOT FIND DEFAULT USERNAME **************" & vbNewLine
            Me.TopMost = False
            MsgBox("User account user was not found on the system", MsgBoxStyle.Critical)
            System.Threading.Thread.Sleep(20000)
        End Try
    End Sub
    Public Sub shield()
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\Screen Shield\Settings", True)
        key.SetValue("PCNo", ComNum.ToString)

    End Sub
    Public Sub changecomname()
        Dim strNewName As String

        strNewName = ComName
        SetComputerName(strNewName)
        TextBox1.Text += "Computername set to " + strNewName
    End Sub
    Public Function getipaddress() As String
        Dim host As Net.IPHostEntry = Net.Dns.GetHostByName(Net.Dns.GetHostName)
        If host.AddressList.Length > 0 Then
            Return host.AddressList(0).ToString
        End If
    End Function
    Public Sub reboot()
        WindowsController.ExitWindows(RestartOptions.Reboot, False)
    End Sub
    Public Sub reg()
        Try
            Dim cool As New Process
            Dim share As New Process
            share.StartInfo.FileName = ("c:\windows\system32\net.exe")
            share.StartInfo.Arguments = ("use s: \\manager\share\reg /user:com1 1")
            cool.StartInfo.FileName = ("c:\windows\system32\regedt32.exe")
            cool.StartInfo.Arguments = ("/s s:\" & ComNum & ".reg")
            cool.Start()
            share.Start()
        Catch ex As Exception

            MsgBox("Can not find Reg file")
            MsgBox(ex)

        End Try





    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = False
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", True)
        key.SetValue("AltDefaultUserName", ComName.ToString)
        key.SetValue("DefaultUserName", ComName.ToString)
        key.SetValue("Shell", "Explorer.exe")
        TextBox1.Text += "Explorer set as shell" & vbNewLine
        reboot()
    End Sub

End Class
