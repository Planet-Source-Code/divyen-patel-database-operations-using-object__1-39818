VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DataBase Operation ( Demo )"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4950
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   24
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add new"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   20
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   ">> |"
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   19
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   18
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   17
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   "| <<"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   16
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   7
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   2160
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   13
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   11
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hire Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Birth Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title Of Courtesy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New DATABASE_OP_CLASS
Private Sub cmd_move_Click(Index As Integer)

If Index = 0 Then
        If c.GET_RECORDSET.BOF <> True Then
            c.GET_RECORDSET.MoveFirst
        End If
ElseIf Index = 1 Then
        If c.GET_RECORDSET.BOF <> True Then
                    c.GET_RECORDSET.MovePrevious
                    If c.GET_RECORDSET.BOF = True Then
                        c.GET_RECORDSET.MoveFirst
                    End If
        End If
ElseIf Index = 2 Then
        If c.GET_RECORDSET.EOF <> True Then
                c.GET_RECORDSET.MoveNext
                If c.GET_RECORDSET.EOF = True Then
                    c.GET_RECORDSET.MoveLast
                End If
        End If
        
ElseIf Index = 3 Then
        If c.GET_RECORDSET.EOF <> True Then
            c.GET_RECORDSET.MoveLast
        End If
End If
FILL_DATA
End Sub



Private Sub Command1_Click(Index As Integer)
Dim st As Boolean
If Index = 0 Then
    CLEAR_TEXT
    c.ADDNEW
    OP_STATUS (False)
ElseIf Index = 1 Then
    For I = 0 To Text1.Count - 1
        c.SET_CURRENT_RECORD I, Text1(I).Text
    Next
    
    
    st = c.UPDATE()
    
    
    If st = False Then
        MsgBox c.errstring, vbInformation, "Error : Update"
        
        c.errstring = ""
        Exit Sub
    End If
    OP_STATUS (True)
    
ElseIf Index = 2 Then
    
    st = c.DELETE()
    If st = False Then
        c.CANCEL_UPDATE
        MsgBox c.errstring, vbInformation, "Error : Delete"
        c.errstring = ""
    End If
    
ElseIf Index = 3 Then
    c.CANCEL_UPDATE
    FILL_DATA
    OP_STATUS (True)
ElseIf Index = 4 Then
    OP_STATUS (False)
End If

End Sub

Private Sub Form_Load()
c.OPEN_CONNECTION "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\NWIND.MDB;Persist Security Info=False"
c.OPEN_RECORD_SET "SELECT * FROM Employees", c.GET_ACTIVE_CONNECTION, adOpenKeyset, adLockOptimistic
FILL_DATA
OP_STATUS (True)
End Sub

Public Function FILL_DATA()
If c.GET_RECORDSET.EOF <> True Then
        For I = 0 To Text1.Count - 1
            Text1(I).Text = c.GET_CURRENT_RECORD(I)
        Next
End If

End Function

Public Function CLEAR_TEXT()
        For I = 0 To Text1.Count - 1
            Text1(I).Text = ""
        Next
End Function

Public Function OP_STATUS(STATUS As Boolean)
    If STATUS = True Then
        Command1(0).Enabled = True
        Command1(1).Enabled = False
        Command1(2).Enabled = True
        Command1(3).Enabled = False
        Command1(4).Enabled = True
        For I = 0 To cmd_move.Count - 1
            cmd_move(I).Enabled = True
        Next
        
        For I = 0 To Text1.Count - 1
            Text1(I).Enabled = False
        Next
        
        
    Else
        Command1(0).Enabled = False
        Command1(1).Enabled = True
        Command1(2).Enabled = False
        Command1(3).Enabled = True
        Command1(4).Enabled = False
        For I = 0 To cmd_move.Count - 1
            cmd_move(I).Enabled = False
        Next
        For I = 0 To Text1.Count - 1
            Text1(I).Enabled = True
        Next
    End If
    
End Function
