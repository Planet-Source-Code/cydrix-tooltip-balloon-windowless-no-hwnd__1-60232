VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balloon Tooltips"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   3390
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   285
      Left            =   1980
      TabIndex        =   7
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   705
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Global Declaration for control without HWND
Dim tipLabel1 As New clsToolTip
Dim tipLabel2 As New clsToolTip

'You can declare them globally too
Dim tipCommand1 As New clsToolTip
Dim tipDrive1 As New clsToolTip

Private Sub Form_Load()
  ' Control with HWND can be declared locally
  Dim tipPicture1 As New clsToolTip
  Dim tipText1 As New clsToolTip
  Dim tipOption1 As New clsToolTip
  Dim tipCheck1 As New clsToolTip
  
  ' Fake that Command1 doesn't have HWND
  tipCommand1.CreateBalloon Me.Command1, 0, "Hey There User!", szBalloon, False, "WOW!", etiInfo, 225, 0
  
  ' Controls without HWND
  tipLabel1.CreateBalloon Me.Label1, 0, "Hey There User!", szBalloon, False, "WOW! Working for Label1", etiError
  tipLabel2.CreateBalloon Me.Label2, 0, "Hey There User!", szBalloon, False, "WOW! Working for Label2", etiWarning
  
  ' normal controls
  tipPicture1.CreateBalloon Me.Picture1, Me.Picture1.hWnd, "Hey There User!", szBalloon, False, "WOW!", etiInfo
  tipDrive1.CreateBalloon Me.Drive1, Me.Drive1.hWnd, "Hey There User!", szBalloon, False, "WOW!", etiInfo
  tipText1.CreateBalloon Me.Text1, Me.Text1.hWnd, "Hey There User!", szBalloon, False, "WOW!", etiInfo
  tipOption1.CreateBalloon Me.Option1, Me.Option1.hWnd, "Hey There User!", szBalloon, False, "WOW!", etiInfo
  tipCheck1.CreateBalloon Me.Check1, Me.Check1.hWnd, "Hey There User!", szBalloon, False, "WOW!", etiInfo
  
End Sub

' we have to set the HWND by calling SetHandle for control without HWND
' class will find the control under the mouse pointer - simple and working !
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tipCommand1.SetHandle Command1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tipLabel1.SetHandle Label1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tipLabel2.SetHandle Label2
End Sub

