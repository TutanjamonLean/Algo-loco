VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23190
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   23190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   14160
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   14160
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Usuarios ingresados"
      Height          =   975
      Left            =   4200
      TabIndex        =   8
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar usuario"
      Height          =   1095
      Left            =   4200
      TabIndex        =   7
      Top             =   5880
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   6000
      TabIndex        =   6
      Top             =   5880
      Width           =   10695
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Text            =   "LOLCITO"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Text            =   "25"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Text            =   "Alcachofa"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Text            =   "Juanito"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar Datos"
      Height          =   1215
      Left            =   6000
      TabIndex        =   0
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label2 
      Height          =   1335
      Left            =   16680
      TabIndex        =   11
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   1215
      Left            =   16920
      TabIndex        =   5
      Top             =   5880
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nombre, apellido, edad, curso As String

Private Sub Command1_Click()
    
    nombre = Text1(0)
    apellido = Text1(1)
    edad = Text1(2)
    curso = Text1(3)
    
'    Label1.Caption = "Nombre: " & Text1(0) & vbCrLf & "Apellido: " & Text1(1) & vbCrLf & "Edad: " & Text1(2) & vbCrLf & "Curso: " & Text1(3)
    
    
End Sub

Private Sub Command2_Click()
    
    List1.AddItem "El alumno " & nombre & " " & apellido & " que tiene " & edad & " años de edad, esta en el curso de " & curso
    
End Sub

Private Sub Command3_Click()
    
'    List1.Clear

    Label1.Caption = ""
    
    Label1.Caption = "Se cargo la cantidad de " & List1.ListCount & " usuarios"
    
'    Print List1.ListCount
'
'    Print List1.ListIndex

'    Print List1.SelCount
    
End Sub

Private Sub Command4_Click()
    
    List1.ListIndex
    
End Sub

Private Sub List1_Click()
    
    Label2.Caption = List1.ListIndex + 1
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
'-----------------------------------------NOMBRE------------------------------------------------
        Case Is = 0
            If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Then
                
                KeyAscii = KeyAscii
                
            ElseIf KeyAscii = 8 Then
                
                KeyAscii = KeyAscii
                
            ElseIf KeyAscii = 32 Then
                
                KeyAscii = KeyAscii
                
            Else
            
                KeyAscii = 0
                
            End If
'------------------------------------------APELLIDO---------------------------------------------------
        Case Is = 1
            If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Then
                
                KeyAscii = KeyAscii
                
            ElseIf KeyAscii = 8 Then
                
                KeyAscii = KeyAscii
            
            ElseIf KeyAscii = 32 Then
                
                KeyAscii = KeyAscii
                
            Else
            
                KeyAscii = 0
                
            End If
'-------------------------------------------EDAD----------------------------------------------
        Case Is = 2
            If KeyAscii >= 48 And KeyAscii <= 57 Then
                
                KeyAscii = KeyAscii
                
            ElseIf KeyAscii = 8 Then
                
                KeyAscii = KeyAscii
                
            Else
                
                KeyAscii = 0
                
            End If
'---------------------------------------------CURSO-------------------------------------
        Case Is = 3
        
            If KeyAscii >= 65 And KeyAscii <= 90 Then
                
                KeyAscii = KeyAscii
                
            ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
                
                KeyAscii = KeyAscii - 32
                
            ElseIf KeyAscii = 8 Then
                
                KeyAscii = KeyAscii
                
            ElseIf KeyAscii = 32 Then
                
                KeyAscii = KeyAscii
                
            Else
                
                KeyAscii = 0
                
            End If
            
        
        
    End Select
    
End Sub
