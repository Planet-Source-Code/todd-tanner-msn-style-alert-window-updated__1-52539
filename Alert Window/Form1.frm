VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Alert"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Hello"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Alert"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    Dim Alert As frmAlert
    Set Alert = New frmAlert
    Caption = Alert.DisplayAlert(Text1.Text)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
