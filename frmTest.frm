VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6132
   ClientLeft      =   8868
   ClientTop       =   5148
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6132
   ScaleWidth      =   6540
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   408
      Left            =   4932
      TabIndex        =   12
      Top             =   5436
      Width           =   372
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   408
      Left            =   4464
      TabIndex        =   11
      Top             =   5436
      Width           =   372
   End
   Begin Proyecto1.WinShape WinShape6 
      Height          =   1056
      Left            =   4608
      TabIndex        =   10
      Top             =   2844
      Width           =   1344
      _ExtentX        =   1990
      _ExtentY        =   1545
      BackColor       =   -2147483646
      Shape           =   4
      FillColor       =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Anti-Alias"
      Height          =   336
      Left            =   1512
      TabIndex        =   9
      Top             =   5436
      Width           =   2532
   End
   Begin Proyecto1.WinShape WinShape5 
      Height          =   876
      Left            =   1044
      TabIndex        =   7
      Top             =   4068
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   1545
      BorderColor     =   255
      Shape           =   2
      BorderWidth     =   10
   End
   Begin VB.TextBox Text2 
      Height          =   480
      Left            =   396
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   4284
      Width           =   912
   End
   Begin Proyecto1.WinShape WinShape4 
      Height          =   912
      Left            =   3492
      TabIndex        =   6
      Top             =   2664
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   1609
      BorderColor     =   16711680
      Shape           =   3
      FillColor       =   65280
      FillStyle       =   0
      BorderWidth     =   3
   End
   Begin VB.PictureBox Picture1 
      Height          =   1200
      Left            =   3708
      ScaleHeight     =   1152
      ScaleWidth      =   1512
      TabIndex        =   5
      Top             =   3132
      Width           =   1560
   End
   Begin Proyecto1.WinShape WinShape3 
      Height          =   1344
      Left            =   1188
      TabIndex        =   4
      Top             =   2340
      Width           =   768
      _ExtentX        =   2053
      _ExtentY        =   910
      BorderColor     =   16711688
      Shape           =   2
      FillStyle       =   7
   End
   Begin Proyecto1.WinShape WinShape1 
      Height          =   1164
      Left            =   1116
      TabIndex        =   1
      Top             =   720
      Width           =   1416
      _ExtentX        =   2498
      _ExtentY        =   2053
      BorderColor     =   255
      Shape           =   3
      BorderWidth     =   10
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   660
      Left            =   864
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmTest.frx":0000
      Top             =   2628
      Width           =   1416
   End
   Begin Proyecto1.WinShape WinShape2 
      Height          =   1164
      Left            =   3492
      TabIndex        =   2
      Top             =   432
      Width           =   1668
      _ExtentX        =   2498
      _ExtentY        =   2053
      BorderColor     =   255
      Shape           =   1
      FillStyle       =   0
      BorderWidth     =   10
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   732
      Left            =   684
      TabIndex        =   0
      Top             =   612
      Width           =   4728
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "WinShape" Then
            iCtl.AntiAlias = Check1.Value = 1
        End If
    Next
End Sub

Private Sub Command2_Click()
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "WinShape" Then
            iCtl.Width = iCtl.Width + 50
            iCtl.Height = iCtl.Height + 50
        End If
    Next
End Sub

Private Sub Command3_Click()
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "WinShape" Then
            iCtl.Width = iCtl.Width - 50
            iCtl.Height = iCtl.Height - 50
        End If
    Next
End Sub

Private Sub Form_Load()
    Check1_Click
End Sub
