VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ֶ�ѡƷ�����ʼ��� v0.5f Python�Ż��濪����......"
   ClientHeight    =   3720
   ClientLeft      =   10380
   ClientTop       =   5715
   ClientWidth     =   7125
   Icon            =   "�ֶ�ѡƷ�����ʼ���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "������"
      Height          =   2655
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   2895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ԥ��100$��˰������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   2220
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ԥ��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   840
         TabIndex        =   17
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   2055
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "webforex.hermes.hexun.com"
      URL             =   "http://webforex.hermes.hexun.com/forex/quotelist?code=FOREXUSDCNY&column=Price"
      Document        =   "/forex/quotelist?code=FOREXUSDCNY&column=Price"
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������Ӧ��ֵ"
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   3735
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ֵ����ۿ���"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "���ҽ��������վ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   960
         MousePointer    =   10  'Up Arrow
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�·��۸������ԪΪ��λ!"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buff������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Steam�󹺼�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " �������ۿ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����鿴"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����(Enter)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(��������)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   750
      TabIndex        =   12
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ʵʱ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text4 = "" Or Text1 = "" Or Text2 = "" Then
MsgBox "��������Ӧ��ֵ���ٽ��м��㣡", vbOKOnly
Else
    Dim E2!, B2!, A2!, D2!, G2!, H2!, sum!, J2 As Single
    E2 = Text3 * 100: B2 = Text2: A2 = Text1: D2 = E2 * Text4
    G2 = E2 / B2 * A2 - D2
    sum = G2 - (G2 * 0.035)
    Label3.Caption = sum
    J2 = sum / D2
        If J2 > 0 Then
        Label4.ForeColor = RGB(255, 0, 0)
        ElseIf J2 <= 0 Then
        Label4.ForeColor = RGB(0, 235, 12)
        End If
    Label4.Caption = Format(J2 * 100, "0.###") & "%"
End If
End Sub



Private Sub Form_Load()
Dim asd As String, bot As String, i As Integer
asd = Inet1.OpenURL(asd)
For i = 1 To Len(asd)
If Mid(asd, i, 1) >= 0 And Mid(asd, i, 1) <= 9 Then
bot = bot & Mid(asd, i, 1)
End If
Next
Text3 = bot / 10000
End Sub

Private Sub Label2_Click()
Dim ab As String, bc As String
ab = Text1: bc = Text2
If ab = "" Then
    MsgBox "������Buff�����ļ۸�", vbOKOnly
ElseIf bc = "" Then
    MsgBox "������Steam�󹺵ļ۸�", vbOKOnly
Else
    Label2.Caption = Format(ab / bc, "0.###")
End If
End Sub

Private Sub Label7_Click()
Shell "explorer https://themoneyconverter.com/CN/CNY/USD", 1
End Sub

