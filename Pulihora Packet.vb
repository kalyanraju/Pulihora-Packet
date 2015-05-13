Option Strict Off
Public Class Form1
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents AxPacketXCtrl1 As AxPacketXLib.AxPacketXCtrl
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.AxPacketXCtrl1 = New AxPacketXLib.AxPacketXCtrl
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.AxPacketXCtrl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.Location = New System.Drawing.Point(0, 27)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(722, 205)
        Me.ListView1.TabIndex = 1
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'AxPacketXCtrl1
        '
        Me.AxPacketXCtrl1.Enabled = True
        Me.AxPacketXCtrl1.Location = New System.Drawing.Point(112, 32)
        Me.AxPacketXCtrl1.Name = "AxPacketXCtrl1"
        Me.AxPacketXCtrl1.OcxState = CType(resources.GetObject("AxPacketXCtrl1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxPacketXCtrl1.Size = New System.Drawing.Size(192, 192)
        Me.AxPacketXCtrl1.TabIndex = 2
        Me.AxPacketXCtrl1.Visible = False
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(301, 1)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Start"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(382, 1)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Stop"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(12, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(283, 21)
        Me.ComboBox1.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(464, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Syn Count :"
        '
        'Form1
        '
        Me.ClientSize = New System.Drawing.Size(722, 233)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.AxPacketXCtrl1)
        Me.Controls.Add(Me.ListView1)
        Me.Name = "Form1"
        Me.Text = "Pulihora Packet"
        CType(Me.AxPacketXCtrl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Dim i As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Add columns to list view
        ListView1.Columns().Add("No.", 100, HorizontalAlignment.Left)
        ListView1.Columns().Add("Date", 100, HorizontalAlignment.Left)
        ListView1.Columns().Add("Protocol", 100, HorizontalAlignment.Left)
        ListView1.Columns().Add("Src IP", 100, HorizontalAlignment.Left)
        ListView1.Columns().Add("Dest IP", 100, HorizontalAlignment.Left)
        ListView1.Columns().Add("Syn Flag", 100, HorizontalAlignment.Left)
        ListView1.Columns().Add("IP id", 100, HorizontalAlignment.Left)

        'Refresh adapters collection
        AxPacketXCtrl1.Reset()

        'Display the default adapter
        For i = 1 To AxPacketXCtrl1.Adapters.Count
            ComboBox1.Items.Add(AxPacketXCtrl1.Adapters(i).Description)
        Next
        'MsgBox(AxPacketXCtrl1.Adapter.Description)

        'Enable menu items
        CaptureStop = False
        CaptureStart = True
    End Sub

    Private Sub Form1_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Stop the running capture 
        If CaptureStop = True Then

            AxPacketXCtrl1.Stop()
        End If
    End Sub

    Private Sub FileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        End
    End Sub



    Private oPacketColl As New PacketXLib.PacketCollectionClass
    Dim vnCounter As Integer

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        'Display packet data
        If ListView1.SelectedItems.Count > 0 Then
            Dim oPacket As PacketXLib.Packet
            oPacket = oPacketColl(ListView1.SelectedIndices().Item(0) + 1)
            Dim vByte As Byte
            Dim sData As String
            Dim nPosition, nColumns As Integer
            nColumns = 16
            For Each vByte In oPacket.DataArray
                If nPosition = 8 Then
                    sData = sData + " "
                End If
                If nPosition >= nColumns Then
                    sData = sData + vbCrLf : nPosition = 1
                Else
                    nPosition = nPosition + 1
                End If
                If vByte <= &HF Then
                    sData = sData + "0"
                End If
                sData = sData + Hex(vByte) + " "
            Next
            MsgBox(sData, MsgBoxStyle.Information, "Packet Detail")
        End If
    End Sub

    Private Sub AxPacketXCtrl1_OnPacket(ByVal sender As Object, ByVal e As AxPacketXLib._IPktXPacketXCtrlEvents_OnPacketEvent) Handles AxPacketXCtrl1.OnPacket
        'Display packet date and size
        Dim I As Short
        Dim thisPacket As String
        Dim SourceIP As String = ""
        Dim DestIP As String = ""
        Dim prtcl As String = ""
        Dim Ipid As String = e.pPacket.Data(19).ToString("X") & "," & e.pPacket.Data(20).ToString("X")
        If e.pPacket.Data(14) = 69 And e.pPacket.Data(23) = 6 Then
            SourceIP = e.pPacket.Data(26) & "." & _
            e.pPacket.Data(27) & "." & + _
            e.pPacket.Data(28) & "." & + _
            e.pPacket.Data(29)
            DestIP = e.pPacket.Data(30) & "." & _
            e.pPacket.Data(31) & "." & + _
            e.pPacket.Data(32) & "." & + _
            e.pPacket.Data(33)
            thisPacket = ""
            For I = 0 To e.pPacket.DataSize - 1
                thisPacket = thisPacket & Chr(e.pPacket.Data(I))
            Next

            If e.pPacket.Protocol = PacketXLib.PktXProtocolType.PktXProtocolTypeTCP Then
                prtcl = "TCP"
            End If
            If e.pPacket.Protocol = PacketXLib.PktXProtocolType.PktXProtocolTypeUDP Then
                prtcl = "UDP"
            End If
            If e.pPacket.Protocol = PacketXLib.PktXProtocolType.PktXProtocolTypeIP Then
                prtcl = "IP"
            End If
            Dim aItem As New ListViewItem(CStr(vnCounter))
            aItem.SubItems().Add(CStr(e.pPacket.Date))
            aItem.SubItems().Add(prtcl)
            aItem.SubItems().Add(SourceIP)
            aItem.SubItems().Add(DestIP)
            aItem.SubItems().Add("1")
            aItem.SubItems().Add(Ipid)
            ListView1.Items().Add(aItem)
            'Save packet
            oPacketColl.Add(e.pPacket)
            vnCounter += 1
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Dim CaptureStart, CaptureStop As Boolean
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Clear list view
        ListView1.Items().Clear()
        'Clear packet collection
        oPacketColl.RemoveAll()
        vnCounter = 0
        'Start capture
        AxPacketXCtrl1.Adapter.BPFilter = "tcp[13] & (0x02|0x04) != 0"
        AxPacketXCtrl1.Start()
        CaptureStart = False
        CaptureStop = True
        Button3.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Stop capture
        AxPacketXCtrl1.Stop()
        CaptureStart = True
        CaptureStop = False
        Button2.Enabled = True
        Button3.Enabled = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        AxPacketXCtrl1.Adapter = AxPacketXCtrl1.Adapters(ComboBox1.SelectedIndex + 1)
        Button2.Enabled = True

    End Sub
End Class
