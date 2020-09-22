VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code For Beginners"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Count the length of String"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Count a specified character"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Count Words"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Count Vowels"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Default Sentence"
      Top             =   480
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please type some text in this box"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox CountVowels(Text1.Text, True) & " vowels are found!", vbInformation
End Sub

Private Sub Command2_Click()
MsgBox UBound(Split(Text1.Text, " ")) + 1 & " words Found !", vbInformation
End Sub

Private Sub Command3_Click()
If Trim(Text2.Text) = "" Then
MsgBox "Please enter a character to search", vbInformation
Exit Sub
End If
MsgBox UBound(Split(LCase(Text1.Text), LCase(Text2.Text))) & " Times found!", vbInformation
End Sub

Private Sub Command4_Click()
MsgBox "Length of the String is : " & Len(Text1.Text), vbInformation
End Sub
Function CountVowels(Text As String, IncludeY As Boolean) As Integer
    Dim X As Integer
    Dim VowelCount As Integer
    Dim CurrLett As String


    For X = 1 To Len(Text)
        CurrLett = Mid$(Text, X, 1)
        CurrLett = UCase$(CurrLett)


        If CurrLett = "A" Or CurrLett = "E" Or CurrLett = "I" Or CurrLett = "O" Or CurrLett = "U" Or (CurrLett = "Y" And IncludeY = True) Then
            VowelCount = VowelCount + 1
        End If
    Next X
    CountVowels = VowelCount
End Function
