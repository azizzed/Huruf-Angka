VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdInputData 
      Caption         =   "Input Data"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInputData_Click()
a = InputBox("Masukkan Input:")
For i = 1 To Len(a)
d = Mid(a, i, 1)
If (d >= "A" And d <= "Z") Or (d >= "a" And d <= "z") Then
Debug.Print "Huruf"
Else
Debug.Print "Bukan Huruf"
End If
Next i
End Sub

Private Sub cmdKeluar_Click()
 If (MsgBox("Yakin Anda Ingin Keluar?", vbYesNo, "Keluar") = vbYes) Then
        Unload Me
    End If
End Sub
