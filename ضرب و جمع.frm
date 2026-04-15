VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' ------ Õ”«» «·„Ã„Ê⁄ „‰ 1 ≈·Ï 100 ------
    Dim sum As Long '  ’ÕÌÕ: Long »œ· Low0
    Dim i As Integer
    sum = 0 '  ÂÌ∆… «·„ €Ì—
    
    For i = 1 To 100
        sum = sum + i
    Next i
    
    MsgBox "„Ã„Ê⁄ «·√⁄œ«œ „‰ 1 ≈·Ï 100 ÂÊ: " & sum

    ' ------ Õ”«» Õ«’· «·÷—» „‰ 1 ≈·Ï 50 ------
    Dim product As Double
    product = 1 ' «· ÂÌ∆… »ﬁÌ„… 1 ·√‰ «·÷—» Ì»œ√ „‰ 1
    
    For i = 1 To 50
        product = product * i
        List1.AddItem "Õ«’· ÷—» " & i & " = " & product '  ’ÕÌÕ: List1.AddItem
        
    Next i
End Sub
