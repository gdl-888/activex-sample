VERSION 5.00
Begin VB.UserControl NumberPad 
   BackStyle       =   0  '투명
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ScaleHeight     =   2595
   ScaleWidth      =   1830
   Begin VB.CommandButton cmdClear 
      Caption         =   "*"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "#"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   9
      Left            =   720
      TabIndex        =   10
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdWordBlock 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblPassword 
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "NumberPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim PW As String
Dim nums(9) As Integer

Public Event SetPassword(pinval)

Private Sub cmdClear_Click()
    PW = ""
    lblPassword.Caption = ""
End Sub

Private Sub cmdSubmit_Click()
    RaiseEvent SetPassword(PW)
End Sub

'http://vb-helper.com/howto_randomize_array.html 참고함
Private Sub cmdWordBlock_Click(Index As Integer)
    PW = PW & cmdWordBlock(Index).Caption
    lblPassword.Caption = lblPassword.Caption & "*"

    Dim i As Integer
    Dim j As Integer
    Dim Temp As Integer

    Randomize Timer
    For i = 0 To 9
        j = Int(Rnd * (9 - i) + i)

        Temp = nums(i)
        nums(i) = nums(j)
        nums(j) = Temp
    Next i
    
    For i = 0 To 9
        cmdWordBlock(i).Caption = nums(i)
    Next i
End Sub

Private Sub UserControl_Initialize()
    PW = ""
    Dim i As Integer
    For i = 0 To 9
        nums(i) = i
        cmdWordBlock(i).Caption = i
    Next i
End Sub
