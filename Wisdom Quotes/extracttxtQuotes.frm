VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quotes"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "extracttxtQuotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin DHTMLEDLibCtl.DHTMLEdit d1 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5055
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   5055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Select"
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4440
      TabIndex        =   8
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Author :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0/0"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "By :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Welcome
'This is a Quotes Program Which Have Above 5500 Quotes From Various Persons .
' The Program is Created in 3 Hours After Fetching The Website WisdomQuotes.com
'Here You Will Find Some Functions That Will Load Quotes From Text File Database And Saperate By Authors

'Author : Ajay Kumar @ AjayWares
'Email : ajay_bnl@yahoo.com


Private Sub Combo1_Change()
Combo1_Click
End Sub

'Load Quotes
Private Sub Combo1_Click()
'Load Quotes ( By Categorie or author)
LoadList Combo1.Text
If UBound(Entrys) > 0 Then
HScroll1.Max = UBound(Entrys)
HScroll1.Min = 1
HScroll1.Value = 1
Label2.Caption = HScroll1.Value & "/" & UBound(Entrys) & " of " & TotalEntrys
HScroll1_Change
End If

End Sub


Private Sub Form_Load()

LoadCates
If Combo1.ListCount > 0 Then
Combo1.ListIndex = 0
End If
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "By : " & Entrys(HScroll1.Value).Author
d1.DocumentHTML = "<font size='2' color='blue'>" & Entrys(HScroll1.Value).Quote & "</font>"
Label2.Caption = HScroll1.Value & "/" & UBound(Entrys) & " of " & TotalEntrys

End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub Label4_Click()
MsgBox "A Little Software Created By AjayWares " & vbCrLf & "Author : Ajay Kumar" & vbCrLf & "Email : ajay_bnl@yahoo.com" & vbCrLf & "Ripped Website : WisdomQuotes.com" & vbCrLf & "Only For Personal / Distribution Use ."

End Sub


