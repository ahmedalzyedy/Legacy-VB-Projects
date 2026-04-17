VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17715
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   17715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ÎÑæÌ"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   1200
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   1680
      ScaleHeight     =   2595
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      Height          =   2295
      Left            =   6600
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      Height          =   2055
      Left            =   6480
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   10080
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shapes As Variant
Dim CurrentShape As Integer

Private Sub Form_Load()
    ' ŞÇÆãÉ ÇáÃÔßÇá ÇáãÑÇÏ ÚÑÖåÇ (ÃÓãÇÁ ÚäÇÕÑ ÇáÜ Shape)
    Shapes = Array("Shape1", "Shape2", "Shape3")
    CurrentShape = 0
    
    ' ÅÙåÇÑ ÇáÔßá ÇáÃæá æÅÎİÇÁ ÇáÈÇŞí
    Controls(Shapes(CurrentShape)).Visible = True
    Timer1.Enabled = True
End Sub




Private Sub Timer1_Timer()
    ' ÅÎİÇÁ ÇáÔßá ÇáÍÇáí
    Controls(Shapes(CurrentShape)).Visible = False
    
    ' ÇáÇäÊŞÇá Åáì ÇáÔßá ÇáÊÇáí
    CurrentShape = CurrentShape + 1
    If CurrentShape > UBound(Shapes) Then
        CurrentShape = 0 ' ÅÚÇÏÉ ãä ÇáÈÏÇíÉ
    End If
    
    ' ÅÙåÇÑ ÇáÔßá ÇáÌÏíÏ
    Controls(Shapes(CurrentShape)).Visible = True
End Sub

Private Sub Command1_Click()
    Unload Me 'end
    
End Sub

