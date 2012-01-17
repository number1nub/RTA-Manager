VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} protect 
   Caption         =   "Enter password..."
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "protect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "protect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CommandButton2_Click()
    OK = False: Unload Me
End Sub


Sub CommandButton1_Click()
    If TextBox1 = sMyPassWord Then
        OK = True
        Unload Me
    Else
        Label1.Caption = "Incorrect password. Try again."
        TextBox1 = ""
        Label1.ForeColor = &HC0&
        TextBox1.SelStart = 0
    End If
End Sub


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
            If TextBox1 = sMyPassWord Then
        OK = True
        Unload Me
    Else
        Label1.Caption = "Incorrect password. Try again."
        TextBox1 = ""
        TextBox1.ForeColor = &HC0&
        
    End If
    End If
End Sub

Private Sub UserForm_Initialize()
    Application.EnableCancelKey = xlErrorHandler
    TextBox1.ForeColor = &H0&
End Sub
