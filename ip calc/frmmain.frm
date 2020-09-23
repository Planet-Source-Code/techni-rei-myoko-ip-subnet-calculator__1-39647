VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subnetting Practice"
   ClientHeight    =   8664
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6120
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8664
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save to a file"
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstmain 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Tag             =   "10"
      ToolTipText     =   "This can pass 255 and start over again"
      Top             =   5160
      Width           =   5895
      _ExtentX        =   10393
      _ExtentY        =   4255
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Subnet"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Network Address"
         Object.Width           =   2461
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Hosts"
         Object.Width           =   3916
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Broadcast Address"
         Object.Width           =   2671
      EndProperty
   End
   Begin VB.ComboBox cbomain 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "This number will represent ""n"""
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtmain 
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Draw out the first ten(10) subnets for this network:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   4920
      Width           =   5895
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "borrowed to subnet the network?"
      Height          =   255
      Index           =   16
      Left            =   600
      TabIndex        =   33
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   4320
      TabIndex        =   32
      ToolTipText     =   "Take the NW portions and change them to 255, then make the n leftmost bits set to 1"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "What is the maximum number of bits that can be "
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   31
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   30
      ToolTipText     =   "Obtained by: uS * uH"
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   4320
      TabIndex        =   29
      ToolTipText     =   "Obtained by: 2^(M-n)-2  (This number will represent ""uH"")"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   28
      ToolTipText     =   "Obtained by: (2^n)-2 (This number will represent ""uS"")"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   27
      ToolTipText     =   "Obtained by: S * H"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   26
      ToolTipText     =   "Obtained by: 2^(M-n) (This number will represent ""H"")"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   25
      ToolTipText     =   "Obtained by: 2^n (This number will represent ""S"")"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   24
      ToolTipText     =   "A: 22 B: 14 CDE: 6 (This number will represent ""M"")"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   23
      ToolTipText     =   "A: NW.255.255.255 B: NW.NW.255.255 CDE: NW.NW.NW.255 (NW is the Network address octets)"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   22
      ToolTipText     =   "A: 255.0.0.0 B: 255.255.0.0. CDE: 255.255.255.0"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   21
      ToolTipText     =   "A: 1 B: 2 CDE: 3"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   20
      ToolTipText     =   "A: 3 B: 2 CDE: 1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblansw 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      ToolTipText     =   "A: 0-127 B: 128-191 C: 192-223 D: 224-239 E: 240-255"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculate the subnet mask for this network:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Total number of usable hosts:"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   17
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Total number of usable hosts per subnet:"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   16
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Total number of usable subnets:"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Total number of hosts:"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Total number of hosts per subnet:"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Total number of subnets:"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "bits to subnet the network and provide the following information"
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   11
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "Borrow"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "What is the broadcast address for this network?"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "What is the default mask for this network?"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "How many octets are dedicated to the networks?"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "How many octets are dedicated to the hosts?"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "What class is this address?"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblmain 
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim spoth As Integer, file As Integer
Dim ishtml As Boolean
Public Function bool2byte(value As String) As Byte
Dim temp As Byte
temp = 0
If Mid(value, 1, 1) = "1" Then temp = temp + 128
If Mid(value, 2, 1) = "1" Then temp = temp + 64
If Mid(value, 3, 1) = "1" Then temp = temp + 32
If Mid(value, 4, 1) = "1" Then temp = temp + 16
If Mid(value, 5, 1) = "1" Then temp = temp + 8
If Mid(value, 6, 1) = "1" Then temp = temp + 4
If Mid(value, 7, 1) = "1" Then temp = temp + 2
If Mid(value, 8, 1) = "1" Then temp = temp + 1
bool2byte = temp
End Function
Public Sub cbomain_Click()
Dim maxbits As Byte
maxbits = (Val(lblansw(5)) + 2)
If cbomain.Text <> "0" Then
    lblansw(6) = 2 ^ Val(cbomain.Text)
    lblansw(7) = 2 ^ (maxbits - Val(cbomain.Text))
    lblansw(8) = Val(lblansw(6)) * Val(lblansw(7))
    lblansw(9) = lblansw(6) - 2
    lblansw(10) = lblansw(7) - 2
    lblansw(11) = Val(lblansw(9)) * Val(lblansw(10))
Else
    lblansw(6) = 0
    lblansw(7) = 2 ^ maxbits
    lblansw(8) = lblansw(7)
    lblansw(9) = 0
    lblansw(10) = lblansw(7) - 2
    lblansw(11) = lblansw(10)
End If

Dim tempb As String
tempb = String(Val(cbomain.Text), "1") & String(maxbits, "0")
Dim bits(1 To 3)
For spoth = 1 To Int(maxbits / 8)
    bits(spoth) = bool2byte(Mid(tempb, ((spoth - 1) * 8) + 1, 8))
Next
Select Case lblansw(0)
    Case ""
    Case "A"
        lblansw(12) = "255." & bits(1) & "." & bits(2) & "." & bits(3)
    Case "B"
        lblansw(12) = "255.255." & bits(1) & "." & bits(2)
    Case Else
        lblansw(12) = "255.255.255." & bits(1)
End Select

lstmain.ListItems.Clear
For spoth = 1 To Val(lstmain.Tag)
    addsubnet
Next

Dim spot1, spot2
For spot1 = 1 To lstmain.ColumnHeaders.count
    If spot1 = 1 Then
        For spot2 = 1 To lstmain.ListItems.count
            If Me.TextWidth(lstmain.ListItems.Item(spot2).Text) + (6 * 15) > lstmain.ColumnHeaders.Item(1).width Then
                lstmain.ColumnHeaders.Item(spot1).width = Me.TextWidth(lstmain.ListItems.Item(spot2).Text) + (6 * 15)
            End If
        Next
    Else
        For spot2 = 1 To lstmain.ListItems.count
            If Me.TextWidth(lstmain.ListItems.Item(spot2).SubItems(spot1 - 1)) + (6 * 15) > lstmain.ColumnHeaders.Item(spot1).width Then
                lstmain.ColumnHeaders.Item(spot1).width = Me.TextWidth(lstmain.ListItems.Item(spot2).SubItems(spot1 - 1)) + (6 * 15)
            End If
        Next
    End If
Next
End Sub

Private Sub cbomain_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Public Function writetofile(asTEXT As String, asHTML As String, Optional asBOTH As String)
    If ishtml = True Then writetofile = asHTML & asBOTH Else writetofile = asTEXT & asBOTH
End Function
Public Function formatted(width As Integer, lefts As String, rights As String, Optional delimeter As String) As String
    If delimeter = Empty Then delimeter = " "
    If Len(lefts) > width Then MsgBox Len(lefts)
    formatted = lefts & String(width - Len(lefts) - Len(rights), delimeter) & rights
End Function
Public Sub SaveFile(filename As String)
On Error GoTo error
Const width As Integer = 100
file = FreeFile
Open filename For Output As file
    Print #file, writetofile(Me.Caption & vbNewLine & String(width, "-"), "<HTML><HEAD><TITLE>" & Me.Caption & "</TITLE><H1>" & Me.Caption & "</H1><HR></HEAD><BODY><CENTER><TABLE BORDER=2>")
    Print #file, writetofile(formatted(width, lblmain(0).Caption, txtmain.Text), "<TR><TD><CENTER><B>" & lblmain(0).Caption & "</TD><TD><CENTER>" & txtmain & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(1).Caption, lblansw(0).Caption), "<TR><TD>" & lblmain(1).Caption & "</TD><TD><CENTER>" & lblansw(0).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(2).Caption, lblansw(1).Caption), "<TR><TD>" & lblmain(2).Caption & "</TD><TD><CENTER>" & lblansw(1).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(3).Caption, lblansw(2).Caption), "<TR><TD>" & lblmain(3).Caption & "</TD><TD><CENTER>" & lblansw(2).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(4).Caption, lblansw(3).Caption), "<TR><TD>" & lblmain(4).Caption & "</TD><TD><CENTER>" & lblansw(3).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(5).Caption, lblansw(4).Caption), "<TR><TD>" & lblmain(5).Caption & "</TD><TD><CENTER>" & lblansw(4).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(15).Caption & lblmain(16).Caption, lblansw(5).Caption), "<TR><TD>" & lblmain(15).Caption & " " & lblmain(16).Caption & "</TD><TD><CENTER>" & lblansw(5).Caption & "</TD></TR>")
    Print #file, writetofile(vbNewLine, "<TR><TD COLSPAN=2><center><B>", lblmain(6).Caption & " " & cbomain.Text & " " & lblmain(7).Caption) & writetofile("", "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(8).Caption, lblansw(6).Caption), "<TR><TD>" & lblmain(8).Caption & "</TD><TD><CENTER>" & lblansw(6).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(9).Caption, lblansw(7).Caption), "<TR><TD>" & lblmain(9).Caption & "</TD><TD><CENTER>" & lblansw(7).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(10).Caption, lblansw(8).Caption), "<TR><TD>" & lblmain(10).Caption & "</TD><TD><CENTER>" & lblansw(8).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(11).Caption, lblansw(9).Caption), "<TR><TD>" & lblmain(11).Caption & "</TD><TD><CENTER>" & lblansw(9).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(12).Caption, lblansw(10).Caption), "<TR><TD>" & lblmain(12).Caption & "</TD><TD><CENTER>" & lblansw(10).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(13).Caption, lblansw(11).Caption), "<TR><TD>" & lblmain(13).Caption & "</TD><TD><CENTER>" & lblansw(11).Caption & "</TD></TR>")
    Print #file, writetofile(formatted(width, lblmain(14).Caption, lblansw(12).Caption), "<TR><TD>" & lblmain(14).Caption & "</TD><TD><CENTER>" & lblansw(12).Caption & "</TD></TR>")
    Print #file, writetofile(vbNewLine & formatted(width, "Subnet    Network Address    Hosts", "Broadcast Address"), "</TABLE><P><TABLE BORDER=2><TR><TH><CENTER>Subnet</TH><TH><CENTER>Network Address</TH><TH><CENTER>Hosts</TH><TH><CENTER>Broadcast Address</TH></TR>")
    Dim temp As String
    If lstmain.ListItems.count > 0 Then
        For spoth = 1 To lstmain.ListItems.count
            With lstmain.ListItems
            temp = writetofile(formatted(Len("Subnet    "), .Item(spoth).Text, ""), "<TR><TD><CENTER>" & .Item(spoth).Text & "</TD>")
            temp = temp & writetofile(formatted(Len("Network Address    "), .Item(spoth).SubItems(1), ""), "<TD><CENTER>" & .Item(spoth).SubItems(1) & "</TD>")
            temp = temp & writetofile(formatted(width - Len(.Item(spoth).SubItems(3)) - Len(temp), .Item(spoth).SubItems(2), ""), "<TD><CENTER>" & .Item(spoth).SubItems(2) & "</TD>")
            temp = temp & writetofile(.Item(spoth).SubItems(3), "<TD><CENTER>" & .Item(spoth).SubItems(3) & "</TD></TR>")
            Print #file, temp
            End With
        Next
    End If
    Print #file, writetofile("", "</TABLE></BODY></HTML>")
Close file
If MsgBox("The file is complete, would you like to see it now?", vbYesNo, "Open generated file?") = vbYes Then Shell "start """ & filename & """", vbHide
error:
If Err.number <> 0 Then Call MsgBox("An error occured, I was unable to save the file", vbCritical, "Could not save")
End Sub

Private Sub cmdsave_Click()
Dim filename As String, extention As String
InitSaveDlg
filename = Save_File(Me.hWnd)
If countchars(filename, ".") = 1 Then filename = filename & ".html"
If filename <> Empty And filename <> ".html" Then
    extention = LCase(getword(filename, countwords(filename, "."), "."))
    If extention <> "html" And extention <> "htm" And extention <> "txt" Then
        Call MsgBox("I'm sorry but I can't save to that format (" & extention & ")", vbCritical, "Can not save. (Unknown format)")
    Else
        If extention = "txt" Then ishtml = False Else ishtml = True
        SaveFile (filename)
    End If
End If
End Sub

Public Sub Form_Click()
Static hasrun As Boolean
If hasrun = False Then
    hasrun = True
    Call MsgBox("Thank you for using my IP Subnet calculator", vbMsgBoxRtlReading, "Thank you")
End If
End Sub

Private Sub Form_Load()
cmdsave.Picture = Me.Icon
End Sub

Private Sub lblansw_Click(Index As Integer)
Form_Click
End Sub

Private Sub lblmain_Click(Index As Integer)
If Index <> 17 Then Form_Click Else Call lstmain_MouseUp(vbLeftButton, 0, 0, 0)
End Sub
Public Function mult(inputnum As Integer) As Integer
    If inputnum > 10 And inputnum < 8192 Then mult = inputnum * 2
    If inputnum = 10 Then mult = 256
    If inputnum = 8192 Then mult = 10
End Function
Public Function divi(inputnum As Integer) As Integer
    If inputnum > 256 Then divi = Int(inputnum / 2):
    If inputnum = 10 Then divi = 8192
    If inputnum = 256 Then divi = 10
End Function
Public Function spellitout(number As Integer) As String
    If number = 0 Then spellitout = "zero"
    If number = 10 Then spellitout = "ten"
    If number = 256 Then spellitout = "two hundred fifty six"
    If number = 512 Then spellitout = "five hundred twelve"
    If number = 1024 Then spellitout = "one thousand twenty four"
    If number = 2048 Then spellitout = "two thousand forty eight"
    If number = 4096 Then spellitout = "four thousand ninety six"
    If number = 8192 Then spellitout = "eight thousand one hundred ninety two"
End Function

Private Sub lstmain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
lstmain.Tag = mult(lstmain.Tag)
Else
lstmain.Tag = divi(lstmain.Tag)
End If
lblmain(17).Caption = "Draw out the first " & spellitout(lstmain.Tag) & "(" & lstmain.Tag & ") subnets for this network:"
cbomain_Click
End Sub

Private Sub txtmain_Change()
Dim isanip As Boolean
isanip = False

    If countwords(txtmain, ".") = 4 Then
        isanip = True
        For spoth = 0 To 3
            If IsNumeric(getword(txtmain, spoth, ".")) = False Then
                isanip = False
            Else
                If Val(getword(txtmain, spoth, ".")) > 255 Then isanip = False
            End If
        Next
    End If
    
    For spoth = 0 To lblansw.UBound
        lblansw(0) = Empty
    Next
    
    If isanip = True Then
        docalc
    End If
    
End Sub
Public Sub answers(host As Byte, nw As String, default As String, broadcast As String)
    Dim maxbits As Byte
    maxbits = host * 8 - 2
    lblansw(1) = host
    lblansw(2) = nw
    lblansw(3) = default
    lblansw(4) = broadcast
    lblansw(5) = maxbits
    cbomain.Clear
    cbomain.Text = "0"
    For spoth = 0 To maxbits
        cbomain.AddItem spoth
    Next
End Sub
Public Sub docalc()
lblansw(0) = getIPclass(Val(getword(txtmain, 0, ".")))
Select Case lblansw(0)
    Case Empty
        Call answers(0, "0", "0.0.0.0", "255.255.255.255")
    Case "A"
        Call answers(3, "1", "255.0.0.0", getword(txtmain, 0, ".") & ".255.255.255")
    Case "B"
        Call answers(2, "2", "255.255.0.0", getword(txtmain, 0, ".") & "." & getword(txtmain, 1, ".") & ".255.255")
    Case Else
        Call answers(1, "3", "255.255.255.0", getword(txtmain, 0, ".") & "." & getword(txtmain, 1, ".") & "." & getword(txtmain, 2, ".") & ".255")
End Select
cbomain_Click
End Sub
Public Function getIPclass(octet As Integer) As String
    If iswithin(octet, 0, 127) Then getIPclass = "A"
    If iswithin(octet, 128, 191) Then getIPclass = "B"
    If iswithin(octet, 192, 223) Then getIPclass = "C"
    If iswithin(octet, 224, 239) Then getIPclass = "D"
    If iswithin(octet, 240, 255) Then getIPclass = "E"
End Function
Public Function iswithin(var As Integer, min As Byte, max As Byte) As Boolean
    If var >= min And var <= max Then iswithin = True Else iswithin = False
End Function
Public Sub addsubnet()
Const a1 = 16777216
Const a2 = 65536
Const a3 = 256
With lstmain.ListItems
Dim numb As Integer, temp
.Add , , lstmain.ListItems.count
numb = .count

'Network address portion
temp = (numb - 1) * Val(lblansw(7).Caption)
Dim bit3(0 To 2)
bit3(0) = Int(temp / a2)
bit3(1) = Int((temp Mod a2) / a3)
bit3(2) = Int(temp Mod a3)

'Hosts first portion
temp = (numb - 1) * Val(lblansw(7).Caption) + 1
Dim bit4(0 To 2)
bit4(0) = Int(temp / a2)
bit4(1) = Int((temp Mod a2) / a3)
bit4(2) = Int(temp Mod a3)

'Hosts second portion
temp = numb * Val(lblansw(7).Caption) - 2
Dim bit2(0 To 2)
bit2(0) = Int(temp / a2)
bit2(1) = Int((temp Mod a2) / a3)
bit2(2) = Int(temp Mod a3)

'Broadcast address portion
temp = (numb * Val(lblansw(7).Caption)) - 1
Dim bits(0 To 2)
bits(0) = Int(temp / a2)
bits(1) = Int((temp Mod a2) / a3)
bits(2) = Int(temp Mod a3)

Dim word As String
Select Case lblansw(0)
    Case ""
    Case "A"
        word = getword(txtmain, 0, ".") & "."
        .Item(numb).SubItems(3) = word & bits(0) & "." & bits(1) & "." & bits(2)
        .Item(numb).SubItems(1) = word & bit3(0) & "." & bit3(1) & "." & bit3(2)
        .Item(numb).SubItems(2) = word & bit4(0) & "." & bit4(1) & "." & bit4(2) & " - " & word & bit2(0) & bit2(1) & "." & "." & bit2(2)
    Case "B"
        word = getword(txtmain, 0, ".") & "." & getword(txtmain, 1, ".") & "."
        .Item(numb).SubItems(3) = word & bits(1) & "." & bits(2)
        .Item(numb).SubItems(1) = word & bit3(1) & "." & bit3(2)
        .Item(numb).SubItems(2) = word & bit4(1) & "." & bit4(2) & " - " & word & bit2(1) & "." & bit2(2)
    Case Else
        word = getword(txtmain, 0, ".") & "." & getword(txtmain, 1, ".") & "." & getword(txtmain, 2, ".") & "."
        .Item(numb).SubItems(3) = word & bits(2)
        .Item(numb).SubItems(1) = word & bit3(2)
        .Item(numb).SubItems(2) = word & bit4(2) & " - " & word & bit2(2)
End Select

If bits(0) > 255 Then .Remove .count
If .count > 1 Then
    If .Item(.count).SubItems(2) = .Item(.count - 1).SubItems(2) Then .Remove .count
End If
End With
End Sub
Private Sub txtmain_KeyPress(KeyAscii As Integer)
If countwords(txtmain, ".") > 4 Then KeyAscii = 0
If countwords(txtmain, ".") = 4 And KeyAscii = Asc(".") Then KeyAscii = 0
If KeyAscii <> 8 And KeyAscii <> 23 And KeyAscii <> 3 And KeyAscii <> 24 Then
If iswithin(KeyAscii, Asc("0"), Asc("9")) = False And KeyAscii <> Asc(".") Then KeyAscii = 0
End If

If iswithin(KeyAscii, Asc("0"), Asc("9")) = True Then
    If Val(getword(txtmain, findcurrword(txtmain, txtmain.SelStart, "."), ".") & Chr(KeyAscii)) > 255 Then
        KeyAscii = 0
    End If
End If
If iswithin(KeyAscii, Asc("0"), Asc("9")) = True Then
    If Len(getword(txtmain, findcurrword(txtmain, txtmain.SelStart, "."), ".") & Chr(KeyAscii)) > 3 Then
        KeyAscii = 0
    End If
End If
End Sub
