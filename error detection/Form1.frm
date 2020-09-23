VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Error"
   ClientHeight    =   5280
   ClientLeft      =   5700
   ClientTop       =   1950
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   3810
   Begin VB.CommandButton Command4 
      Caption         =   "Error occure after line numbered 123"
      Height          =   345
      Left            =   300
      TabIndex        =   4
      Top             =   1590
      Width           =   3315
   End
   Begin VB.TextBox Text1 
      Height          =   2505
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   2130
      Width           =   3285
   End
   Begin VB.CommandButton Command3 
      Caption         =   "None Sub ID Stated"
      Height          =   345
      Left            =   300
      TabIndex        =   2
      Top             =   1080
      Width           =   3315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Error Code AtError Occur At SubID 65335"
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   630
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Error Occur At Sub ID 10"
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   210
      Width           =   3315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cstrModuleName = "Form1.frm"

Private Sub Command1_Click()
10
On Error GoTo err1
Dim x As Integer
    x = 12 \ 0
    
Exit Sub
err1:
    MsgBox Err.Number & " - " & Err.Description, , cstrModuleName & " - " & Erl
End Sub

Private Sub Command2_Click()
'The amount of code that can be loaded into a form, class,
'or standard module is limited to 65,534 lines
65535
On Error GoTo err2
Dim x
    x = 12 \ 0
    
Exit Sub
err2:
    MsgBox Err.Number & " - " & Err.Description, , cstrModuleName & " - " & Erl
End Sub

Private Sub Command3_Click()
On Error GoTo err3
Dim x
    x = 12 \ 0
    
Exit Sub
err3:
    MsgBox Err.Number & " - " & Err.Description, , cstrModuleName & " - " & Erl
End Sub

Private Sub Command4_Click()
100 'line #100 defined
On Error GoTo err2
Dim x
123 'line #123 defined
    x = 12 \ 0
    
150 'line #150 defined
Exit Sub
err2:
    MsgBox Err.Number & " - " & Err.Description, , cstrModuleName & " - " & Erl
End Sub

Private Sub Form_Load()
Dim strMsg As String

    strMsg = "For more accuracy, you can state the number "
    strMsg = strMsg & "in what ever line you want except the "
    strMsg = strMsg & "function/sub declaration. "
    strMsg = strMsg & "Erl will take the latest number pass on "
    strMsg = strMsg & "current sub or function before error occurs."
    Text1.Text = strMsg
End Sub
