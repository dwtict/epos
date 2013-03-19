VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BIXOLON Customer Display - Windows Driver Sample"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Reverse"
      Height          =   650
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Macro"
      Height          =   650
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "j"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FontControl Excute"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear 2nd line"
      Height          =   650
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear 1st line"
      Height          =   650
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "BCD 2nd Line Display"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BCD 2nd Line Display"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Screen"
      Height          =   660
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "BCD 1st Line Display"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BCD 1st Line Display"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Printer.Font.Size = 7
Printer.FontName = "FontControl"
Printer.Print "d"
Printer.Print "m"

Printer.Font.Size = 7
Printer.FontName = "BCD 1st Line"

Printer.Print Text1.Text
Printer.EndDoc

End Sub


Private Sub Command2_Click()
Printer.Font.Size = 7
Printer.FontName = "FontControl"

Printer.Print "a"
Printer.EndDoc
End Sub

Private Sub Command3_Click()
Printer.Font.Size = 7
Printer.FontName = "BCD 2nd Line"

Printer.Print Text2.Text
Printer.EndDoc
End Sub

Private Sub Command4_Click()
Printer.Font.Size = 7
Printer.FontName = "FontControl"

Printer.Print "b"
Printer.EndDoc
End Sub

Private Sub Command5_Click()
Printer.Font.Size = 7
Printer.FontName = "FontControl"

Printer.Print "c"
Printer.EndDoc
End Sub

Private Sub Command6_Click()
Printer.Font.Size = 7
Printer.FontName = "FontControl"

Printer.Print Text3.Text
Printer.EndDoc
End Sub

Private Sub Command7_Click()
Printer.Font.Name = "FontControl"
Printer.Print "a"
Printer.Print "b"
Printer.Print "c"
Printer.Print "d"
Printer.Print "n"

Printer.Print "a"
Printer.Print "d"

Printer.Font.Name = "BCD 1st Line"
Printer.Print "   Save your money"

Printer.Font.Name = "BCD 2nd Line"
Printer.Print "    with BIXOLON"

Printer.Font.Name = "FontControl"
Printer.Print "e"
Printer.Print "n"
Printer.Print "o"
Printer.EndDoc
End Sub

Private Sub Command8_Click()
Printer.Font.Size = 7
Printer.FontName = "FontControl"

Printer.Print "l"
Printer.Font.Size = 7
Printer.FontName = "BCD 1st Line"
Printer.Print "Reverse Test"

Printer.FontName = "BCD 2nd Line"
Printer.Print "Reverse Test"
Printer.EndDoc
End Sub

Private Sub Form_Load()
     
For Each prnPrinter In Printers
    If prnPrinter.DeviceName = "BIXOLON BCD-1000" Then
        Set Printer = prnPrinter
        Exit For
    End If
Next

If Printer.DeviceName <> "BIXOLON BCD-1000" Then
    MsgBox ("BIXOLON Customer Display Driver don't installed.")
    Unload Mainform
End If
    
       
End Sub
    

