Imports System.Management
Imports System.Net
Imports System.Net.NetworkInformation
Imports MySql.Data.MySqlClient
Imports System.Data
Imports MySql.Data

Public Class Form1

    'Variable para determinar IP dinamica o estatica: 
    Private niAdpaters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()

    Friend conexion As MySqlConnection

    'Función para determinar IP dinamica o estatica
    Private Function GetDhcp(iSelectedAdpater As Int32) As [Boolean]
        If niAdpaters(iSelectedAdpater).GetIPProperties().GetIPv4Properties() IsNot Nothing Then
            Return niAdpaters(iSelectedAdpater).GetIPProperties().GetIPv4Properties().IsDhcpEnabled
        Else
            Return False
        End If
    End Function


    Private Sub ActualizarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarToolStripMenuItem.Click

        'Sistema Operativo
        Label16.Text = My.Computer.Info.OSFullName

        'RAM
        Dim gb As Double
        Dim b2gb As Double = 1024 * 1024 * 1024
        gb = My.Computer.Info.TotalPhysicalMemory / b2gb
        Label17.Text = String.Format("{0:f2} GB", gb)

        'Processor Name
        Dim ProcessorName As String
        ProcessorName = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CentralProcessor\0", "ProcessorNameString", Nothing)
        Label23.Text = ProcessorName

        'HD Name
        Dim hdgb As Double
        Dim hdgb2 As Double = 1024 * 1024 * 1024
        hdgb = My.Computer.FileSystem.Drives(0).TotalSize / hdgb2
        Label24.Text = String.Format("{0:f2} GB", hdgb)

        'HD Free Space
        Dim hdgbf As Double
        Dim hdgbf2 As Double = 1024 * 1024 * 1024
        hdgbf = My.Computer.FileSystem.Drives(0).AvailableFreeSpace / hdgbf2
        Label25.Text = String.Format("{0:f2} GB", hdgbf)

        'Username
        Label15.Text = SystemInformation.UserName

        'DomainName
        Label32.Text = Environment.UserDomainName

        'IP Status

        'External IP
        Dim wc As New WebClient
        Label22.Text = wc.DownloadString("http://checkip.amazonaws.com/")

        'Ping 8.8.8.8
        If My.Computer.Network.Ping("8.8.8.8") = True Then
            Label29.BackColor = Color.Green
        Else
            Label29.BackColor = Color.Red
        End If

        'Ping Google.com
        If My.Computer.Network.Ping("8.8.8.8") = True Then
            Label39.BackColor = Color.Green
        Else
            Label39.BackColor = Color.Red
        End If

        'Service Tag
        Dim q As New SelectQuery("Win32_bios")
        Dim search As New ManagementObjectSearcher(q)
        Dim info As New ManagementObject
        Dim ST As String

        For Each info In search.Get
            ST = info("serialnumber").ToString
            TextBox1.Text = ST
        Next

        'Manufacturer, Model
        Dim q2 As New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Dim info2 As New ManagementObject

        For Each info2 In q2.Get
            Label30.Text = info2("manufacturer").ToString
            Label31.Text = info2("model").ToString
        Next

        'Computer Name, IPv4, Subnet Mask, Default Gateway, DNS1, DNS2
        'Computer Name
        Label21.Text = System.Net.Dns.GetHostName()
        For Each ip In System.Net.Dns.GetHostEntry(Label21.Text).AddressList
            If ip.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                'IPv4 Adress
                Label26.Text = ip.ToString()

                For Each adapter As Net.NetworkInformation.NetworkInterface In Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                    For Each unicastIPAddressInformation As Net.NetworkInformation.UnicastIPAddressInformation In adapter.GetIPProperties().UnicastAddresses
                        If unicastIPAddressInformation.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                            If ip.Equals(unicastIPAddressInformation.Address) Then
                                'Subnet Mask
                                Label35.Text = unicastIPAddressInformation.IPv4Mask.ToString()

                                Dim adapterProperties As Net.NetworkInformation.IPInterfaceProperties = adapter.GetIPProperties()
                                For Each gateway As Net.NetworkInformation.GatewayIPAddressInformation In adapterProperties.GatewayAddresses
                                    'Default Gateway
                                    Label27.Text = gateway.Address.ToString()
                                Next

                                'DNS1
                                If adapterProperties.DnsAddresses.Count > 0 Then
                                    Label36.Text = adapterProperties.DnsAddresses(0).ToString()
                                End If

                                'DNS2
                                If adapterProperties.DnsAddresses.Count > 1 Then
                                    Label37.Text = adapterProperties.DnsAddresses(1).ToString()
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        Next

        'Ip dinamica o estatica
        If GetDhcp(2) = True Then
            Label38.Text = String.Format("Dinamica")
        Else
            Label38.Text = String.Format("Estatica")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim connstring As String = "server=sharksyst.com;userid=sharksys_Cris;password=*Cris2021*;database=sharksys_Prueba"

        Dim conn As MySqlConnection = Nothing

        Try
            conn = New MySqlConnection(connstring)
            conn.Open()

            Dim query As String = "SELECT * FROM TablaPrueba;"
            Dim da As New MySqlDataAdapter(query, conn)
            Dim ds As New DataSet()
            da.Fill(ds, "TablaPrueba")
            Dim dt As DataTable = ds.Tables("TablaPrueba")

            For Each row As DataRow In dt.Rows
                For Each col As DataColumn In dt.Columns
                    MsgBox(row(col).ToString() + vbTab)
                Next

                MsgBox(vbNewLine)
            Next

        Catch exDB As Exception
            Console.WriteLine("Error: {0}", exDB.ToString())
        Finally
            If conn IsNot Nothing Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim connstring As String = "server=sharksyst.com;userid=sharksys_Cris;password=*Cris2021*;database=sharksys_Prueba"

        Dim conn As MySqlConnection = Nothing

        Try
            conn = New MySqlConnection(connstring)
            conn.Open()

            Dim query As String = "UPDATE TablaPrueba SET nombre='GIH' WHERE id='1';"
            Dim da As New MySqlDataAdapter(query, conn)
            Dim ds As New DataSet()
            da.Fill(ds, "TablaPrueba")
            Dim dt As DataTable = ds.Tables("TablaPrueba")

            For Each row As DataRow In dt.Rows
                For Each col As DataColumn In dt.Columns
                    Console.WriteLine(row(col).ToString() + vbTab)
                Next

                Console.WriteLine(vbNewLine)
            Next

        Catch exDB As Exception
            Console.WriteLine("Error: {0}", exDB.ToString())
        Finally
            If conn IsNot Nothing Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TesteoInternetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TesteoInternetToolStripMenuItem.Click
        'Ping 8.8.8.8 from CMD
        Shell("ping www.google.com -t")
    End Sub

    Private Sub ConfiguraciónIPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfiguraciónIPToolStripMenuItem.Click
        'IPconfig from CMD
        Dim Command = "ipconfig /all"
        Shell("cmd /k" & Command, 1, True)
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'Hipervinculo
        Process.Start("www.soporte.sharksyst.com")
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Outl As Object
        Outl = CreateObject("Outlook.Application")
        If Outl IsNot Nothing Then
            Dim omsg As Object
            omsg = Outl.CreateItem(0)
            omsg.To = "ayuda@sharksyst.com"
            omsg.subject = "Nuevo Ticket - "
            omsg.Display(True) 'will display message to user
        End If

    End Sub
End Class
