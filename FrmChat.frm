VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chatter v.4"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4620
   Icon            =   "FrmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   885
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
      Height          =   345
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   8520
      Width           =   1080
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5730
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Renk deðiþtirmek için çift týklayýn...!!!"
      Top             =   2640
      Width           =   4365
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1395
      Left            =   120
      LinkItem        =   "text3"
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Renk deðiþtirmek için çift týklayýn...!!!"
      Top             =   1200
      Width           =   4365
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2595
      Top             =   1950
   End
   Begin MSWinsockLib.Winsock Winsck 
      Left            =   2310
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   1001
      LocalPort       =   1001
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2055
      Top             =   3105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   315
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   30
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remote IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3135
      TabIndex        =   5
      Top             =   630
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "FrmChat.frx":27A2
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gürkan(R)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   8760
      Width           =   885
   End
   Begin VB.Menu mnufl 
      Caption         =   "File"
      Begin VB.Menu mnue 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnucho 
         Caption         =   "Choose Background"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yanýt As Variant
Private Sub renksec()
        CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
        CommonDialog1.Flags = cdlCCRGBInit
        CommonDialog1.ShowColor
    Form1.BackColor = CommonDialog1.Color
   Exit Sub
ErrHandler:
   Exit Sub
End Sub
Private Sub renksec2()
        CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
        CommonDialog1.Flags = cdlCCRGBInit
        CommonDialog1.ShowColor
    Text2.BackColor = CommonDialog1.Color
    Text3.BackColor = CommonDialog1.Color
   Exit Sub
ErrHandler:
   Exit Sub
End Sub
Private Sub Form_Load()
    Winsck.LocalPort = 1001
    Winsck.RemotePort = 1001
    Label3.Caption = Winsck.LocalIP
End Sub
Private Sub mnucho_Click()
     renksec
End Sub
Private Sub mnue_Click()
    kapat
End Sub
Private Sub Text1_Change()
    Winsck.RemoteHost = Text1.Text
End Sub
Private Sub Text2_Change()
     On Error GoTo err
        Dim Msg As String
        Msg = Text2.Text
        Winsck.SendData Msg
err:
    Select Case err.Number
        Case 10 To 100000
            MsgBox "IP numarasý girmediniz veya numara hatalý", vbCritical, "Hata"
    End Select
End Sub
Private Sub Text2_DblClick()
    renksec2
End Sub
Private Sub Text3_DblClick()
    renksec2
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
        Clear$ = Chr(0)
    If Len(Text2.Text) = 0 And Len(Text1.Text) > 0 Then
        Winsck.SendData Clear$
    End If
    If InStr(1, Text3.Text, Special, vbTextCompare) <> 0 Then
        Winsck.SendData Text2.Text
    End If
End Sub
Private Sub Winsck_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Winsck.GetData Msg, vbString
    Text3.Text = Msg + Special
End Sub
Private Sub Command1_Click()
    kapat
End Sub
Private Sub kapat()
Winsck.Close
    If Text2.Text <> "" Then
            mesaj = "Yapýlan chat kaydedilsin mi?"
            düðme = vbYesNo + vbCritical + vbDefultButton1
            baþlýk = "Kayýt"
            yanýt = MsgBox(mesaj, düðme, baþlýk)
        If yanýt = vbYes Then
             Open "mesaj.log" For Append As #1
                  tarih = Now
                  send = Text2.Text
                  receive = Text3.Text
            Print #1, "Tarih : ", tarih
            Print #1, send
            Print #1, "---------------------"
            Print #1, receive
            Close #1
               MsgBox "kayýt edildi [dosya adý:mesaj.log]", vbOKOnly, "Tamam"
        ElseIf yanýt = vbNo Then
        End If
    Else
End If
    End
End Sub
