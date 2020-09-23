VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   3840
      ScaleHeight     =   1635
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'start the subclassing to trach the mouse enter and mouse exit in the picture box
If Attach(Me.Picture1) = True Then
Me.Print "Subclassing Started"
Else
Me.Print "Subclassing Error"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unhook/Unsubclass Before Exit....
If DeAttach(Me.Picture1) = True Then
Me.Print "Subclassing Started"
Else
Me.Print "Subclassing Error"
End If
End Sub
