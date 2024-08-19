VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   14115
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   3
      Left            =   13680
      TabIndex        =   4
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   2
      Left            =   11040
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   1
      Left            =   8400
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   8400
      TabIndex        =   0
      Top             =   4920
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        
        Case Is = 0
            If KeyAscii >= 65 And KeyAscii <= 90 Then
                
                KeyAscii = KeyAscii
                
            Else
            
                KeyAscii = 0
                
            End If
            
        Case Is = 1
            If KeyAscii >= 65 And KeyAscii <= 90 Then
                
                KeyAscii = KeyAscii
                
            Else
            
                KeyAscii = 0
                
            End If
            
        Case Is = 2
            If KeyAscii <= 48 And KeyAscii >= 57 Then
                
                KeyAscii = KeyAscii
                
            Else
                
                KeyAscii = 0
                
            End If
            
        Case Is = 3
            if
        
        
    End Select
    
End Sub
