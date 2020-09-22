VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Refer to the class module for comments
Option Explicit

Public TestClass As Class1
Private Sub Form_Load()
    Set TestClass = New Class1
    If TestClass.Attach("coozzzzz - coozzzzz") = True Then
        'Successfully attached, change background color and image
        Call TestClass.ChangeBackgroundColor("#000000")
        Call TestClass.ChangeBackgroundImage("http://www.planet-source-code.com/vb/images/PscLogo1.jpg")
    Else
        Set TestClass = Nothing
    End If
End Sub
