
Imports System.Management
Imports System.Net


Class MainWindow

    Private Sub Close_Button_Click(sender As Object, e As RoutedEventArgs) Handles Close_Button.Click
        Me.Close()
    End Sub

    Private Sub Minus_Button_Click(sender As Object, e As RoutedEventArgs) Handles Minus_Button.Click
        Me.Hide()
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim userName = Environment.UserName
        UserSEgg.ToolTip = "您好 :" + Chr(32) + userName

        Dim MyClient As Net.WebClient = New Net.WebClient
        Try
            Dim MyReader As New System.IO.StreamReader(MyClient.OpenRead("https://v1.hitokoto.cn/?c=b&c=a&encode=text"), System.Text.Encoding.UTF8)
            Dim MyWebCode As String = MyReader.ReadToEnd
            Me.Hitokoto.Text = MyWebCode
            MyReader.Close()
        Catch ex As Exception
            Me.Hitokoto.Text = "！- 网络未连接 - ！"
        End Try

    End Sub

    Private Sub Webhome_Click(sender As Object, e As RoutedEventArgs) Handles Webhome.Click
        HomeSnackbarMessage.Content = "官网暂未开放"
        HomeSnackbar.IsActive = True
    End Sub

    Private Sub QQ_Click(sender As Object, e As RoutedEventArgs) Handles QQ.Click
        Process.Start("https://jq.qq.com/?_wv=1027&k=9n37OG8D")
    End Sub

    Private Sub Github_Click(sender As Object, e As RoutedEventArgs) Handles Github.Click
        Process.Start("https://github.com/Forever331")
    End Sub

    Private Sub BiliBIli_Click(sender As Object, e As RoutedEventArgs) Handles BiliBIli.Click
        Process.Start("https://space.bilibili.com/13223538")
    End Sub

    Private Sub CoolApk_Click(sender As Object, e As RoutedEventArgs) Handles CoolApk.Click
        Process.Start("https://www.coolapk.com/u/807260")
    End Sub

    Private Sub HomeSnackbarMessage_ActionClick_1(sender As Object, e As RoutedEventArgs) Handles HomeSnackbarMessage.ActionClick
        HomeSnackbar.IsActive = False
    End Sub

    Private Sub TaskbarIcon_TrayMouseDoubleClick(sender As Object, e As RoutedEventArgs)
        Me.Show()
        Me.WindowState = WindowState.Normal
    End Sub

    Private Sub OpenSYS_Click(sender As Object, e As RoutedEventArgs) Handles OpenSYS.Click
        Me.Show()
        Me.WindowState = WindowState.Normal
    End Sub

    Private Sub ExitSYS_Click(sender As Object, e As RoutedEventArgs) Handles ExitSYS.Click
        Me.Close()
    End Sub

    Private Sub HomeLOGO_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles HomeLOGO.MouseLeftButtonDown
        SP_Info.Visibility = Visibility.Hidden
        SP_Apps.Visibility = Visibility.Hidden
        SP_Main.Visibility = Visibility.Visible

    End Sub

    Private Sub Info_Click(sender As Object, e As RoutedEventArgs) Handles Info.Click
        SP_Main.Visibility = Visibility.Hidden
        SP_Apps.Visibility = Visibility.Hidden
        SP_Info.Visibility = Visibility.Visible

    End Sub

    Private Sub Apps_Click(sender As Object, e As RoutedEventArgs) Handles Apps.Click
        SP_Main.Visibility = Visibility.Hidden
        SP_Info.Visibility = Visibility.Hidden
        SP_Apps.Visibility = Visibility.Visible
    End Sub

    Private Sub Info_Test_Click(sender As Object, e As RoutedEventArgs) Handles Info_Test.Click
        Info_ListBox.Items.Clear()

        Dim OpSystem As New ManagementObjectSearcher("select * from win32_OperatingSystem")
        Dim OpS As ManagementObjectCollection = OpSystem.Get()
        For Each OpSys As Object In OpS
            Info_ListBox.Items.Add("Windows版本:" + Chr(32) + OpSys.GetPropertyValue("Caption").ToString())
            Info_ListBox.Items.Add("系统类型:" + Chr(32) + OpSys.GetPropertyValue("OSArchitecture").ToString() + "操作系统")
        Next

        Dim CmSystem As New ManagementObjectSearcher("select * from win32_computersystem")
        Dim CmS As ManagementObjectCollection = CmSystem.Get()
        For Each CmSys As Object In CmS
            Dim A As Double = (CmSys.GetPropertyValue("TotalPhysicalMemory") / 1024 / 1024 / 1024)
            A = Int(A)
            Info_ListBox.Items.Add("工作组:" + Chr(32) + CmSys.GetPropertyValue("Domain").ToString())
            Info_ListBox.Items.Add("计算机名称:" + Chr(32) + CmSys.GetPropertyValue("__SERVER").ToString())
            Info_ListBox.Items.Add("计算机制造商:" + Chr(32) + CmSys.GetPropertyValue("Manufacturer").ToString())
            Info_ListBox.Items.Add("计算机内存:" + Chr(32) + (A + 1).ToString() + "GB")
        Next

        Dim Process As New ManagementObjectSearcher("select * from win32_Processor")
        Dim CPUPr As ManagementObjectCollection = Process.Get()
        For Each CPU As Object In CPUPr
            Info_ListBox.Items.Add("CPU型号:" + Chr(32) + CPU.GetPropertyValue("Name").ToString() + Chr(32) + CPU.GetPropertyValue("NumberOfCores").ToString() + "核" + CPU.GetPropertyValue("NumberOfLogicalProcessors").ToString() + "线程")
        Next

        Dim NETconfig As New ManagementObjectSearcher("select * from win32_NetworkAdapterConfiguration")
        Dim MNetconfig As ManagementObjectCollection = NETconfig.Get()
        Info_ListBox.Items.Add("IP信息:")
        For Each MNF As ManagementObject In MNetconfig
            If MNF("IPEnabled") Then
                Info_ListBox.Items.Add(Chr(32) + Chr(32) + "内网IP地址:" + Chr(32) + MNF.GetPropertyValue("IpAddress")(0).ToString())
                Info_ListBox.Items.Add(Chr(32) + Chr(32) + "MAC地址:" + Chr(32) + MNF.GetPropertyValue("MACAddress").ToString())
                Exit For
            End If
        Next

        Dim objSearcher As New ManagementObjectSearcher("SELECT * FROM Win32_SoundDevice")
        Dim objCollection As ManagementObjectCollection = objSearcher.Get()
        Info_ListBox.Items.Add("声卡信息:")
        For Each obj As ManagementObject In objCollection
            Info_ListBox.Items.Add(Chr(32) + Chr(32) + obj.GetPropertyValue("Caption").ToString())
        Next

        Dim net As New ManagementObjectSearcher("SELECT *FROM   Win32_NetworkAdapter WHERE Manufacturer != 'Microsoft' AND NOT PNPDeviceID LIKE 'ROOT\\%'")
        Dim ne1 As ManagementObjectCollection = net.Get()
        Info_ListBox.Items.Add("网卡信息:")
        For Each obj As ManagementObject In ne1
            Info_ListBox.Items.Add(Chr(32) + Chr(32) + obj.GetPropertyValue("Name").ToString())
        Next

        Dim VideoCont As New ManagementObjectSearcher("SELECT * FROM Win32_VideoController")
        Dim VideoGet As ManagementObjectCollection = VideoCont.Get()
        Info_ListBox.Items.Add("显卡信息:")
        For Each Video As ManagementObject In VideoGet
            Info_ListBox.Items.Add(Chr(32) + Chr(32) + Video.GetPropertyValue("Name").ToString())
        Next


    End Sub

    Private Sub AppSnackbarMessage_ActionClick(sender As Object, e As RoutedEventArgs)
        AppSnackbar.IsActive = False
    End Sub

    Private Sub AS_SSD_Click(sender As Object, e As RoutedEventArgs) Handles AS_SSD.Click
        Dim AS_SSD As New String("Program\Safe_Disk\AS SSD Benchmark\AS SSD Benchmark.exe")
        If My.Computer.FileSystem.FileExists(AS_SSD) Then
            Process.Start(AS_SSD)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub CrystalDiskInfo_Click(sender As Object, e As RoutedEventArgs) Handles CrystalDiskInfo.Click
        Dim CrystalDiskInfo As New String("Program\Safe_Disk\CrystalDiskInfo\DiskInfo64.exe")
        If My.Computer.FileSystem.FileExists(CrystalDiskInfo) Then
            Process.Start(CrystalDiskInfo)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub CrystalDiskMark_Click(sender As Object, e As RoutedEventArgs) Handles CrystalDiskMark.Click
        Dim CrystalDiskMark As New String("Program\Safe_Disk\CrystalDiskMark\DiskMark64.exe")
        If My.Computer.FileSystem.FileExists(CrystalDiskMark) Then
            Process.Start(CrystalDiskMark)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub HD_Tune_Click(sender As Object, e As RoutedEventArgs) Handles HD_Tune.Click
        Dim HD_Tune As New String("Program\Safe_Disk\HD Tune\HDTunePro 5.5.exe")
        If My.Computer.FileSystem.FileExists(HD_Tune) Then
            Process.Start(HD_Tune)
        Else
            AppSnackbar.IsActive = True
        End If

    End Sub

    Private Sub PartAssist_Click(sender As Object, e As RoutedEventArgs) Handles PartAssist.Click
        Dim PartAssist As New String("Program\Safe_Disk\PartAssist\StartPartAssist64.exe")
        If My.Computer.FileSystem.FileExists(PartAssist) Then
            Process.Start(PartAssist)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub SSD_Z_Click(sender As Object, e As RoutedEventArgs) Handles SSD_Z.Click
        Dim SSD_Z As New String("Program\Safe_Disk\SSD-Z\SSD-Z.exe")
        If My.Computer.FileSystem.FileExists(SSD_Z) Then
            Process.Start(SSD_Z)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub Disk_Benchmark_Click(sender As Object, e As RoutedEventArgs) Handles Disk_Benchmark.Click
        Dim Disk_Benchmark As New String("Program\Safe_Disk\ATTO Disk Benchmark\ATTO 磁盘基准测试.exe")
        If My.Computer.FileSystem.FileExists(Disk_Benchmark) Then
            Process.Start(Disk_Benchmark)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub DiskGenius_Click(sender As Object, e As RoutedEventArgs) Handles DiskGenius.Click
        Dim DiskGenius As New String("Program\Safe_Disk\DiskGenius\DiskGenius(x64).exe")
        If My.Computer.FileSystem.FileExists(DiskGenius) Then
            Process.Start(DiskGenius)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub Victoria_Click(sender As Object, e As RoutedEventArgs) Handles Victoria.Click
        Dim Victoria As New String("Program\Alert_Disk\LLFTOOL\LLFTOOL.exe")
        If My.Computer.FileSystem.FileExists(Victoria) Then
            Process.Start(Victoria)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub LLFTOOL_Click(sender As Object, e As RoutedEventArgs) Handles LLFTOOL.Click
        Dim LLFTOOL As New String("Program\Alert_Disk\Victoria\Victoria.exe")
        If My.Computer.FileSystem.FileExists(LLFTOOL) Then
            Process.Start(LLFTOOL)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub Aida64_Click(sender As Object, e As RoutedEventArgs) Handles Aida64.Click
        Dim Aida64 As New String("Program\Test\Aida64\aida64.exe")
        If My.Computer.FileSystem.FileExists(Aida64) Then
            Process.Start(Aida64)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub CPU_Z_Click(sender As Object, e As RoutedEventArgs) Handles CPU_Z.Click
        Dim CPU_Z As New String("Program\Test\CPU-Z\cpuz_x64.exe")
        If My.Computer.FileSystem.FileExists(CPU_Z) Then
            Process.Start(CPU_Z)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub GPU_Z_Click(sender As Object, e As RoutedEventArgs) Handles GPU_Z.Click
        Dim GPU_Z As New String("Program\Test\GPU-Z\GPU-Z.2.38.0.exe")
        If My.Computer.FileSystem.FileExists(GPU_Z) Then
            Process.Start(GPU_Z)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub HWMonitorPro_Click(sender As Object, e As RoutedEventArgs) Handles HWMonitorPro.Click
        Dim HWMonitorPro As New String("Program\Test\HWMonitorPro\HWMonitorPro_x64.exe")
        If My.Computer.FileSystem.FileExists(HWMonitorPro) Then
            Process.Start(HWMonitorPro)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub HWiNFO_Click(sender As Object, e As RoutedEventArgs) Handles HWiNFO.Click
        Dim HWiNFO As New String("Program\Test\HWiNFO\HWiNFO64.exe")
        If My.Computer.FileSystem.FileExists(HWiNFO) Then
            Process.Start(HWiNFO)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub

    Private Sub DisplayX_Click(sender As Object, e As RoutedEventArgs) Handles DisplayX.Click
        Dim DisplayX As New String("Program\Test\DisplayX\DisplayX.exe")
        If My.Computer.FileSystem.FileExists(DisplayX) Then
            Process.Start(DisplayX)
        Else
            AppSnackbar.IsActive = True
        End If
    End Sub


End Class
