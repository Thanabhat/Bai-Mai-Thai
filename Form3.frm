VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "เกี่ยวกับ ใบไม้ไทย"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   240
      Picture         =   "Form3.frx":0ECA
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   210
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   8
      X2              =   288
      Y1              =   272
      Y2              =   272
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   288
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ใบไม้ไทย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "ใบไม้ไทย" & Chr(13) & Chr(10) & _
                    "Version: 1.2.0 (Revision 4)" & Chr(13) & Chr(10) & _
                    "This program is a freeware."
    Label2.Caption = "Developed by" & Chr(13) & Chr(10) & _
                    "   Kuankidnatta Arunsri" & Chr(13) & Chr(10) & _
                    "   Thanabhat Koomsubha" & Chr(13) & Chr(10) & _
                    "   Pawee Manee-in" & Chr(13) & Chr(10) & _
                    "   Mahidol Wittayanusorn School, Thailand" & Chr(13) & Chr(10) & _
                    "Email : thanabhat_jo@hotmail.com"
    Label3.Caption = "This program is a part of the scientific project " & Chr(13) & Chr(10) & _
                    Chr(34) & "The comparison of the difference of leaf" & Chr(13) & Chr(10) & _
                    " measurement between using Bai Mai Thai program" & Chr(13) & Chr(10) & _
                    " and Area meter A033" & Chr(34) & Chr(13) & Chr(10) & _
                    "Studied at Mahidol Wittayanusorn School, Thailand"
                    
End Sub

