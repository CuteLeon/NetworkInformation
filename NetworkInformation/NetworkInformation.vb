'---------------------------------------------------------------------
'  This file is part of the Microsoft .NET Framework SDK Code Samples.
' 
'  Copyright (C) Microsoft Corporation.  All rights reserved.
' 
'This source code is intended only as a supplement to Microsoft
'Development Tools and/or on-line documentation.  See these other
'materials for detailed information regarding Microsoft code samples.
' 
'THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
'KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'PARTICULAR PURPOSE.
'---------------------------------------------------------------------

Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Globalization


Public Class NetworkInformation

    Dim networkInterfaces() As NetworkInterface
    Dim currentInterface As NetworkInterface

    Delegate Sub NetworkAddressChangedCallback()
    Delegate Sub NetworkAvailabilityCallback( _
        ByVal isNetworkAvailable As Boolean)

    Private Sub NetworkInformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Wire up the NetworkAddressChanged events so we can get notified
        ' when an address change occurs on	any	of the network interfaces.
        ' These changes occur when the interface changes operational
        ' status (up/down) or a new interface is added. 
        '@Leon：坚挺网络可用性变化和网络地址变化事件
        AddHandler System.Net.NetworkInformation.NetworkChange.NetworkAvailabilityChanged, AddressOf Me.networkChange_NetworkAvailabilityChanged
        AddHandler System.Net.NetworkInformation.NetworkChange.NetworkAddressChanged, AddressOf Me.networkChange_NetworkAddressChanged

        ' Populate the	global interfaces container with the list of all
        ' network interfaces.
        '@Leon：获取所有网络接口
        networkInterfaces = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces

        ' Determine if	the	network	is available at	startup.
        '@Leon：获取启动时网络是否可用
        UpdateNetworkAvailability(System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())

        ' Update the information for the network intefaces.
        '@Leon：更新网络接口列表
        UpdateNetworkInformation()
    End Sub

    Private Sub networkChange_NetworkAvailabilityChanged(ByVal sender As Object, ByVal e As NetworkAvailabilityEventArgs)
        '网络可用性变化
        Me.Invoke(
            New NetworkAvailabilityCallback(AddressOf UpdateNetworkAvailability), New Object() {e.IsAvailable})
    End Sub

    Private Sub networkChange_NetworkAddressChanged(ByVal sender As Object, ByVal e As EventArgs)
        '网络地址变化
        Me.Invoke(
            New NetworkAddressChangedCallback(AddressOf UpdateNetworkInformation))
    End Sub

    '更新网络接口列表
    Private Sub UpdateNetworkInformation()
        networkInterfaces = NetworkInterface.GetAllNetworkInterfaces()
        networkInterfacesComboBox.Items.Clear()

        For Each nic As NetworkInterface In networkInterfaces
            networkInterfacesComboBox.Items.Add(nic.Description)
        Next

        If networkInterfaces.Length = 0 Then
            networkInterfacesComboBox.Items.Add( _
                "No NICs found on the machine.")
        Else
            currentInterface = networkInterfaces(0)
            UpdateCurrentNicInformation()
        End If

        networkInterfacesComboBox.SelectedIndex = 0
    End Sub

    '更新currentInterface接口的信息
    Private Sub UpdateCurrentNicInformation()
        ' Set the DNS suffix if any exists
        Dim ipProperties As IPInterfaceProperties = currentInterface.GetIPProperties()
        dnsSuffixTextLabel.Text = ipProperties.DnsSuffix.ToString()

        ' Display the IP address information associated with this
        ' interface including anycast,	unicast, multicast,	DNS	servers,
        ' WINS	servers, DHCP servers, and the gateway
        addressListView.Items.Clear()
        Dim anycastInfo As IPAddressInformationCollection = ipProperties.AnycastAddresses
        For Each info As IPAddressInformation In anycastInfo
            InsertAddress(info.Address, "Anycast")
        Next
        Dim unicastInfo As UnicastIPAddressInformationCollection = ipProperties.UnicastAddresses
        For Each info As UnicastIPAddressInformation In unicastInfo
            InsertAddress(info.Address, "Unicast")
        Next
        Dim multicastInfo As MulticastIPAddressInformationCollection = ipProperties.MulticastAddresses
        For Each info As MulticastIPAddressInformation In multicastInfo
            InsertAddress(info.Address, "Multicast")
        Next
        Dim gatewayInfo As GatewayIPAddressInformationCollection = ipProperties.GatewayAddresses
        For Each info As GatewayIPAddressInformation In gatewayInfo
            InsertAddress(info.Address, "Gateway")
        Next

        Dim ipAddresses As IPAddressCollection = ipProperties.WinsServersAddresses
        InsertAddresses(ipAddresses, "WINS Server")
        ipAddresses = ipProperties.DhcpServerAddresses
        InsertAddresses(ipAddresses, "DHCP Server")
        ipAddresses = ipProperties.DnsAddresses
        InsertAddresses(ipAddresses, "DNS Server")
    End Sub

    Private Sub InsertAddresses(ByVal ipAddresses As IPAddressCollection, ByVal addressType As String)
        For Each address As IPAddress In ipAddresses
            InsertAddress(address, addressType)
        Next
    End Sub

    Private Sub InsertAddress(ByVal address As IPAddress, ByVal addressType As String)
        Dim listViewInformation(2) As String
        listViewInformation(0) = address.ToString()
        listViewInformation(1) = addressType

        Dim item As ListViewItem = New ListViewItem(listViewInformation)
        addressListView.Items.Add(item)
    End Sub

    Private Sub UpdateNetworkAvailability(ByVal isNetworkAvailable As Boolean)
        If isNetworkAvailable Then
            networkAvailabilityTextLabel.Text = "At least one network interface is up."
        Else
            networkAvailabilityTextLabel.Text = "The network is not currently available."
        End If
    End Sub

    Private Sub networkInterfacesComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles networkInterfacesComboBox.SelectedIndexChanged
        currentInterface = networkInterfaces(networkInterfacesComboBox.SelectedIndex)
        UpdateCurrentNicInformation()
    End Sub

    Private Sub updateInfoTimer_Tick( _
        ByVal sender As System.Object, _
        ByVal e As System.EventArgs) _
        Handles updateInfoTimer.Tick
        UpdateNicStats()
    End Sub

    Private Sub UpdateNicStats()
        ' Get the IPv4	statistics for the currently selected interface
        Dim ipStats As IPv4InterfaceStatistics = currentInterface.GetIPv4Statistics()

        Dim numberFormat As NumberFormatInfo = NumberFormatInfo.CurrentInfo
        Dim bytesReceivedInKB As Long = ipStats.BytesReceived / 1024
        Dim bytesSentInKB As Long = ipStats.BytesSent / 1024

        speedTextLabel.Text = GetSpeedString(currentInterface.Speed)
        bytesReceivedTextLabel.Text = bytesReceivedInKB.ToString("N0", numberFormat) + " KB"
        bytesSentTextLabel.Text = bytesSentInKB.ToString("N0", numberFormat) + " KB"

        operationalStatusTextLabel.Text = currentInterface.OperationalStatus.ToString()
        supportsMulticastTextLabel.Text = currentInterface.SupportsMulticast.ToString()
    End Sub

    Private Shared Function GetSpeedString(ByVal speed As Long) As String
        Select Case speed
            Case 10000000
                GetSpeedString = "10 MB"
            Case 11000000
                GetSpeedString = "11 MB"
            Case 54000000
                GetSpeedString = "54 MB"
            Case 100000000
                GetSpeedString = "100 MB"
            Case 1000000000
                GetSpeedString = "1 GB"
            Case Else
                GetSpeedString = speed.ToString(NumberFormatInfo.CurrentInfo)
        End Select

    End Function

End Class
