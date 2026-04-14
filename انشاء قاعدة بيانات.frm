VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   4560
   ClientTop       =   2370
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   12465
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4080
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "«·Œ—ÊÃ"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«Œ— «·”Ã·"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "«·”Ã· «·”«»ﬁ"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«·”Ã· «·À«‰Ì"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "«·”Ã· «·«Ê·"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ–›"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«÷«›…"
      Height          =   495
      Left            =   8640
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "«·⁄‰Ê«‰"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "—ﬁ„ «·„ÊŸ›"
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Data1.Recordset.AddNew
    Me.Text1.SetFocus
End Sub

Private Sub Command2_Click()
    If MsgBox("Â·  —Ìœ «·Õ–›ø", vbYesNo) = vbYes Then
        Data1.Recordset.Delete
    End If
End Sub

Private Sub Command3_Click()
 Data1.Recordset.MoveFirst
End Sub

Private Sub Command4_Click()
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then Data1.Recordset.MoveLast

End Sub

Private Sub Command5_Click()
    Data1.Recordset.MovePrevious
    If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveLast
End Sub


Private Sub Form_Load()
    On Error Resume Next
    Data1.DatabaseName = "C:\visual basic 6\VB98\BIBLIO.MDB"
    Data1.RecordSource = "EMP"
'    Data1.Recordset.Close
    Data1.Refresh
End Sub


' Œ—ÊÃ „⁄  √ﬂÌœ
Private Sub Command7_Click()
    If MsgBox("Â·  —Ìœ «·Œ—ÊÃø", vbYesNo) = vbYes Then
        End
    End If
End Sub
