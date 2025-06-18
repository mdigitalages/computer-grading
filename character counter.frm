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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim phrase As String, char As String
Dim length As Integer
Dim count As Integer, i As Integer

phrase = InputBox("Enter a string: ")
char = InputBox("Enter the character :")

If char = "" Then
 char = ""
End If
Print "Phrase entered is: " & phrase
length = Len(phrase)
Print "length of phrase is: " & length
Print
Print "character entered is: """ & char & """"
Print
For i = 1 To lenght
 If Mid$(phrase, 1, 1) = char Then
  count = count + 1
  End If
Next i

Print "Number of """; char; """ in """; phrase; """ is: "; count
End Sub
