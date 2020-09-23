VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Select The Drives You Want To Hide..."
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR ALL"
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "APPLY"
      Height          =   495
      Left            =   720
      TabIndex        =   26
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CheckBox Check6 
      Caption         =   "F"
      Height          =   495
      Left            =   600
      TabIndex        =   25
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "E"
      Height          =   495
      Left            =   600
      TabIndex        =   24
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "D"
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   2730
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "C"
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "B"
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "A"
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox Check7 
      Caption         =   "G"
      Height          =   495
      Left            =   1980
      TabIndex        =   19
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox Check8 
      Caption         =   "H"
      Height          =   495
      Left            =   1980
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check9 
      Caption         =   "I"
      Height          =   495
      Left            =   1995
      TabIndex        =   17
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CheckBox Check10 
      Caption         =   "J"
      Height          =   495
      Left            =   1995
      TabIndex        =   16
      Top             =   2730
      Width           =   1215
   End
   Begin VB.CheckBox Check11 
      Caption         =   "K"
      Height          =   495
      Left            =   1995
      TabIndex        =   15
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CheckBox Check12 
      Caption         =   "L"
      Height          =   495
      Left            =   1995
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check13 
      Caption         =   "M"
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox Check14 
      Caption         =   "N"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check15 
      Caption         =   "O"
      Height          =   495
      Left            =   3405
      TabIndex        =   11
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CheckBox Check16 
      Caption         =   "P"
      Height          =   495
      Left            =   3405
      TabIndex        =   10
      Top             =   2730
      Width           =   1215
   End
   Begin VB.CheckBox Check17 
      Caption         =   "Q"
      Height          =   495
      Left            =   3405
      TabIndex        =   9
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CheckBox Check18 
      Caption         =   "R"
      Height          =   495
      Left            =   3405
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check19 
      Caption         =   "S"
      Height          =   495
      Left            =   4740
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox Check20 
      Caption         =   "T"
      Height          =   495
      Left            =   4740
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check21 
      Caption         =   "U"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CheckBox Check22 
      Caption         =   "V"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   2730
      Width           =   1215
   End
   Begin VB.CheckBox Check23 
      Caption         =   "W"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CheckBox Check24 
      Caption         =   "X"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check25 
      Caption         =   "Y"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox Check26 
      Caption         =   "Z"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim A As Variant


hi = MsgBox("are you sure", vbExclamation + vbOKCancel, "Confirmation")
If (hi = vbCancel) Then
   Exit Sub
Else
' exit button on key logger tab


A = 0
If Check1.value = 1 Then A = A + 1
If Check2.value = 1 Then A = A + 2
If Check3.value = 1 Then A = A + 4
If Check4.value = 1 Then A = A + 8
If Check5.value = 1 Then A = A + 16
If Check6.value = 1 Then A = A + 32
If Check7.value = 1 Then A = A + 64
If Check8.value = 1 Then A = A + 128
If Check9.value = 1 Then A = A + 256
If Check10.value = 1 Then A = A + 512
If Check11.value = 1 Then A = A + 1024
If Check12.value = 1 Then A = A + 2048
If Check13.value = 1 Then A = A + 4096
If Check14.value = 1 Then A = A + 8192
If Check15.value = 1 Then A = A + 16384
If Check16.value = 1 Then A = A + 32768
If Check17.value = 1 Then A = A + 65536
If Check18.value = 1 Then A = A + 131072
If Check19.value = 1 Then A = A + 262144
If Check20.value = 1 Then A = A + 524288
If Check21.value = 1 Then A = A + 1048576
If Check22.value = 1 Then A = A + 2097152
If Check23.value = 1 Then A = A + 4194304
If Check24.value = 1 Then A = A + 8388608
If Check25.value = 1 Then A = A + 16777216
If Check26.value = 1 Then A = A + 33554432





CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", REG_DWORD, A



  
End If
End Sub



Private Sub Command2_Click()
hi = MsgBox("are you sure", vbExclamation + vbOKCancel, "Confirmation")
If (hi = vbCancel) Then
   Exit Sub
Else
' exit button on key logger tab
 Check1.value = 0
 Check2.value = 0
 Check3.value = 0
 Check4.value = 0
 Check5.value = 0
 Check6.value = 0
 Check7.value = 0
 Check8.value = 0
 Check9.value = 0
 Check10.value = 0
 Check11.value = 0
 Check12.value = 0
 Check13.value = 0
 Check14.value = 0
 Check15.value = 0
 Check16.value = 0
 Check17.value = 0
 Check18.value = 0
 Check19.value = 0
 Check20.value = 0
 Check21.value = 0
 Check22.value = 0
 Check23.value = 0
 Check24.value = 0
 Check25.value = 0
 Check26.value = 0
 

  
End If



End Sub

Private Sub Command3_Click()
hi = MsgBox("are you sure", vbExclamation + vbOKCancel, "Confirmation")
If (hi = vbCancel) Then
   Exit Sub
Else
' exit button on key logger tab

Form2.Hide

Form1.Show
  
End If


End Sub
