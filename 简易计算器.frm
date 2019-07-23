VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "简易计算器V1.0 By Youyou475"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5385
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check4 
      Caption         =   "除法"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "乘法"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "减法"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "加法"
      Height          =   300
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Text            =   "计算结果将在此处显示"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "第二个数字："
      BeginProperty Font 
         Name            =   "等线"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "第一个数字："
      BeginProperty Font 
         Name            =   "等线"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Let Check2.Enabled = False
        Let Check3.Enabled = False
        Let Check4.Enabled = False
    End If
    If Check1.Value = 0 Then
        Let Check2.Enabled = True
        Let Check3.Enabled = True
        Let Check4.Enabled = True
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Let Check1.Enabled = False
        Let Check3.Enabled = False
        Let Check4.Enabled = False
    End If
    If Check2.Value = 0 Then
        Let Check1.Enabled = True
        Let Check3.Enabled = True
        Let Check4.Enabled = True
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Let Check2.Enabled = False
        Let Check1.Enabled = False
        Let Check4.Enabled = False
    End If
    If Check3.Value = 0 Then
        Let Check2.Enabled = True
        Let Check1.Enabled = True
        Let Check4.Enabled = True
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Let Check2.Enabled = False
        Let Check3.Enabled = False
        Let Check1.Enabled = False
    End If
    If Check4.Value = 0 Then
        Let Check2.Enabled = True
        Let Check3.Enabled = True
        Let Check1.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
    If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 Then
        Let Text3.Text = "请选择计算类型！"
    End If
    If Check1.Value = 1 Then
        Let a = Text1.Text * 1 + Text2.Text * 1
        Let Text3.Text = a
    End If
    If Check2.Value = 1 Then
        Let a = Text1.Text * 1 - Text2.Text * 1
        Let Text3.Text = a
    End If
    If Check3.Value = 1 Then
        Let a = Text1.Text * Text2.Text
        Let Text3.Text = a
    End If
    If Check4.Value = 1 Then
        Let a = Text1.Text / Text2.Text
        Let Text3.Text = a
    End If
End Sub
