VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "ãºäÁéä·Â 1.2"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   695
   ScaleMode       =   0  'User
   ScaleWidth      =   1060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Switch to original image"
      Height          =   975
      Left            =   1920
      TabIndex        =   47
      Top             =   7800
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3720
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   15
      TabIndex        =   45
      Top             =   9960
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   15
      TabIndex        =   37
      Top             =   6480
      Width           =   3555
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   420
         Left            =   2520
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   120
         TabIndex        =   42
         Text            =   "0"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate"
         Height          =   975
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   420
         Left            =   960
         Max             =   256
         TabIndex        =   39
         Top             =   720
         Value           =   100
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   420
         Left            =   120
         MaxLength       =   256
         TabIndex        =   38
         Text            =   "100"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         Caption         =   "Color depth"
         Height          =   420
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Leaf Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   900
         Left            =   90
         TabIndex        =   41
         Top             =   2400
         Width           =   3330
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Line option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15
      TabIndex        =   33
      Top             =   5400
      Width           =   3555
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   1200
         TabIndex        =   35
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Length"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3720
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   32
      Top             =   3000
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Change image size (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15
      TabIndex        =   29
      Top             =   4320
      Width           =   3555
      Begin VB.HScrollBar HScroll1 
         Height          =   420
         Left            =   840
         Max             =   100
         TabIndex        =   31
         Top             =   480
         Value           =   100
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   420
         Left            =   120
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "100"
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Position and Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   3555
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Height          =   615
         Left            =   1440
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y ="
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X = "
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   25
         Left            =   2640
         TabIndex        =   26
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   24
         Left            =   2040
         TabIndex        =   25
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   23
         Left            =   1440
         TabIndex        =   24
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   22
         Left            =   840
         TabIndex        =   23
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   21
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   20
         Left            =   2640
         TabIndex        =   21
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   19
         Left            =   2040
         TabIndex        =   20
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   18
         Left            =   1440
         TabIndex        =   19
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   17
         Left            =   840
         TabIndex        =   18
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   16
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   15
         Left            =   2640
         TabIndex        =   16
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   14
         Left            =   2040
         TabIndex        =   15
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   13
         Left            =   1440
         TabIndex        =   14
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   12
         Left            =   840
         TabIndex        =   13
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   11
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   10
         Left            =   1440
         TabIndex        =   11
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   9
         Left            =   2040
         TabIndex        =   10
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   2640
         TabIndex        =   9
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   840
         TabIndex        =   8
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   2640
         TabIndex        =   6
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   2040
         TabIndex        =   5
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      DragIcon        =   "Form1.frx":0ECA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3720
      MouseIcon       =   "Form1.frx":11D4
      MousePointer    =   99  'Custom
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Visible         =   0   'False
         X1              =   96
         X2              =   136
         Y1              =   64
         Y2              =   104
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu reset 
         Caption         =   "&Reset image"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu switch 
         Caption         =   "&Switch image..."
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu calculate 
         Caption         =   "&Calculate..."
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu help2 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fname As String
Dim bSelecting As Boolean
Dim bPicture As Boolean
Dim len1 As Double
Dim len2 As Double
Dim total As Double
Dim LX As Double
Dim LY As Double
Dim R As Long
Dim G As Long
Dim B As Long
Dim cut As Long
Dim cnt As Long
Dim color As Long
Dim avg As Long
Dim pic As Long
Const DefaultCut As Long = 100

Private Sub about_Click()
    Form3.Show
End Sub

Private Sub calculate_Click()
    Call Command1_Click
End Sub

Private Sub close_Click()
    For I = 1 To 25
        Label1(I).BackColor = &H80000013
    Next I
    
    bPicture = False
    Call picReset
    Call AllReset
    Call F1DIS
    Call F2DIS
    Call F3DIS
    Call F4DIS
    
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture1.Picture = LoadPicture("")
    Picture2.Picture = LoadPicture("")
    Picture3.Picture = LoadPicture("")
    
    reset.Enabled = False
    
End Sub

Private Sub Combo1_Click()
    Text5.Text = "Sq." & Combo1.Text
End Sub

Private Sub Command1_Click()
    Call picReset
    Call F1DIS
    Call F2DIS
    Call F3DIS
    Call F4DIS
    
    For I = 0 To 100
    Next I
    
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture2.Picture, _
        Picture1.ScaleLeft, Picture1.ScaleTop, _
            Picture1.ScaleWidth, Picture1.ScaleHeight, _
        Picture2.ScaleLeft, Picture2.ScaleTop, _
            Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = Picture1.Image

    ProgressBar1.Value = 0
    Dim pgmd As Long
    pgmd = Picture1.Width \ 100
    
    cnt = 0
    For X = 0 To Picture1.Width
        For Y = 0 To Picture1.Height
            color = Picture1.Point(X, Y)
            R = color Mod 256
            G = (color \ 256) Mod 256
            B = (color \ 256 \ 256) Mod 256
            avg = (R + G + B) \ 3
            If avg > cut Then
                Picture1.PSet (X, Y), RGB(255, 255, 255)
            Else
                Picture1.PSet (X, Y), RGB(0, 0, 0)
                cnt = cnt + 1
            End If
        Next Y
        If X Mod pgmd = 0 And ProgressBar1.Value < 100 Then
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
    Next X
    
    len1 = Val(Text2.Text)
    len2 = Sqr(((Line1.X1 - Line1.X2) ^ 2) + ((Line1.Y1 - Line1.Y2) ^ 2))
    total = (cnt * ((len1 / len2) ^ 2))
    Text4.Text = total
    ProgressBar1.Value = 0

    Call F1EN
    Call F2EN
    Call F3EN
    Call F4EN
End Sub

Private Sub Command2_Click()
    If pic = 1 Then
        Picture3.Height = Picture1.Height
        Picture3.Width = Picture1.Width
        Picture3.AutoRedraw = True
        Picture3.PaintPicture Picture2.Picture, _
            Picture3.ScaleLeft, Picture3.ScaleTop, _
                Picture3.ScaleWidth, Picture3.ScaleHeight, _
            Picture2.ScaleLeft, Picture2.ScaleTop, _
                Picture2.ScaleWidth, Picture2.ScaleHeight
        Picture3.Picture = Picture3.Image
        
        Picture1.Visible = False
        Picture3.Visible = True
        pic = 3
        Command2.Caption = "Switch to calculated image"
    Else
        Picture1.Visible = True
        Picture3.Visible = False
        pic = 1
        Command2.Caption = "Switch to original image"
    End If
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Combo1.AddItem ("mm")
    Combo1.AddItem ("cm")
    Combo1.AddItem ("m")
    Combo1.ListIndex = 0
    Call Combo1_Click
        
    Picture3.Top = Picture1.Top
    Picture3.Left = Picture1.Left
        
    bPicture = False
    Call F1DIS
    Call F2DIS
    Call F3DIS
    Call F4DIS
    Call AllReset
    
    reset.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure tou want to exit?", vbOKCancel + vbQuestion + vbDefaultButton2, Me.Caption) = vbCancel Then
        Cancel = 1
    Else
        End
    End If
End Sub

Private Sub help2_Click()
    Form2.Show
End Sub

Private Sub HScroll1_Change()
    Call picReset
    Call F3DIS
    Call F4DIS
    Line1.Visible = False
    
    Text1.Text = HScroll1.Value

    If bPicture = True Then
        Picture1.Width = Val(Text1.Text) * (Picture2.Width / 100)
        Picture1.Height = Val(Text1.Text) * (Picture2.Height / 100)
        
        Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture2.Picture, _
            Picture1.ScaleLeft, Picture1.ScaleTop, _
                Picture1.ScaleWidth, Picture1.ScaleHeight, _
            Picture2.ScaleLeft, Picture2.ScaleTop, _
                Picture2.ScaleWidth, Picture2.ScaleHeight
        Picture1.Picture = Picture1.Image
    Else
        Picture1.Visible = False
    End If
End Sub

Private Sub HScroll2_Change()
    Text3.Text = HScroll2.Value
    cut = HScroll2.Value
End Sub

Private Sub open_Click()
    CommonDialog1.Filter = "All Pictures" + _
        "|*.bmp;*.dib;*.jpg;*.jpeg;*.gif;*.wmf;" + _
        "*.emf;*.ico;*.cur" + _
        "|Bitmaps (*.bmp;*.dib)|*.bmp;*.dib" + _
        "|GIF Images (*.gif)|*.gif" + _
        "|JPEG Images (*.jpg;*.jpeg)|*.jpg;*.jpeg" + _
        "|Metafiles (*.wmf;*.emf)|*.wmf;*.emf" + _
        "|Icons (*.ico;*.cur)|*.ico;*.cur" + _
        "|All Files (*.*)|*.*"
        
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    fname = CommonDialog1.FileName
    If fname <> "" Then
        bPicture = True
        Call picReset
    
        Picture1.Visible = True
        Picture1.AutoSize = True
        Picture1.Picture = LoadPicture(fname)
        Picture1.AutoSize = False
        
        Picture2.AutoSize = True
        Picture2.Picture = LoadPicture(fname)
        Picture2.AutoSize = False
        
        Line1.Visible = False
        
        Call F1EN
        Call F2EN
        Call AllReset
        
        reset.Enabled = True
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Line1.X1 = X
    Line1.Y1 = Y
    Line1.X2 = X
    Line1.Y2 = Y
    Line1.Visible = True
    bSelecting = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label2.Caption = "X = " & X
    Label3.Caption = "Y = " & Y
    Call ZoomPointer(Picture1, X, Y)
    
    If bSelecting Then
        Line1.X2 = X
        Line1.Y2 = Y
    End If
End Sub

Private Sub ZoomPointer(Picture1 As Object, nX As Single, nY As Single)
    On Error GoTo En

    Label1(1).BackColor = Picture1.Point(nX - 2, nY - 2)
    Label1(2).BackColor = Picture1.Point(nX - 1, nY - 2)
    Label1(3).BackColor = Picture1.Point(nX, nY - 2)
    Label1(4).BackColor = Picture1.Point(nX + 1, nY - 2)
    Label1(5).BackColor = Picture1.Point(nX + 2, nY - 2)
    
    Label1(6).BackColor = Picture1.Point(nX - 2, nY - 1)
    Label1(7).BackColor = Picture1.Point(nX - 1, nY - 1)
    Label1(8).BackColor = Picture1.Point(nX, nY - 1)
    Label1(9).BackColor = Picture1.Point(nX + 1, nY - 1)
    Label1(10).BackColor = Picture1.Point(nX + 2, nY - 1)
    
    Label1(11).BackColor = Picture1.Point(nX - 2, nY)
    Label1(12).BackColor = Picture1.Point(nX - 1, nY)
    Label1(13).BackColor = Picture1.Point(nX, nY)
    Label1(14).BackColor = Picture1.Point(nX + 1, nY)
    Label1(15).BackColor = Picture1.Point(nX + 2, nY)
    
    Label1(16).BackColor = Picture1.Point(nX - 2, nY + 1)
    Label1(17).BackColor = Picture1.Point(nX - 1, nY + 1)
    Label1(18).BackColor = Picture1.Point(nX, nY + 1)
    Label1(19).BackColor = Picture1.Point(nX + 1, nY + 1)
    Label1(20).BackColor = Picture1.Point(nX + 2, nY + 1)
    
    Label1(21).BackColor = Picture1.Point(nX - 2, nY + 2)
    Label1(22).BackColor = Picture1.Point(nX - 1, nY + 2)
    Label1(23).BackColor = Picture1.Point(nX, nY + 2)
    Label1(24).BackColor = Picture1.Point(nX + 1, nY + 2)
    Label1(25).BackColor = Picture1.Point(nX + 2, nY + 2)
    Exit Sub
En:
    Resume Next
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bSelecting And (Line1.X1 <> Line1.X2 Or Line1.Y1 <> Line1.Y2) Then
        Call F3EN
        Call F4EN
         
        len1 = 0
        len1 = Val(InputBox("Input length : ", Me.Caption))
        Text2.Text = len1
    End If
    bSelecting = False
End Sub

Private Sub reset_Click()
    
    Picture1.Visible = True
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(fname)
    Picture1.AutoSize = False
        
    Picture2.AutoSize = True
    Picture2.Picture = LoadPicture(fname)
    Picture2.AutoSize = False
        
    Line1.Visible = False
    
    Call AllReset
    Call picReset
    Call F1EN
    Call F2EN
    Call F3DIS
    Call F4DIS
End Sub

Private Sub switch_Click()
    Call Command2_Click
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Val(Text1.Text) < 0 Then Text1.Text = 0
    If Val(Text1.Text) > 100 Then Text1.Text = 100
    HScroll1.Value = Text1.Text
    
    Call HScroll1_Change
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    nKeyAscii = KeyAscii
    If KeyAscii = 13 Then
        Call Text1_Validate(True)
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    nKeyAscii = KeyAscii
    If KeyAscii = 13 Then
        If MsgBox("Change length to " & Text2.Text & " " & Combo1.Text, vbYesNoCancel + vbQuestion, Me.Caption) = vbYes Then
            len1 = Val(Text2.Text)
        Else
            Text2.Text = len1
        End If
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If Val(Text3.Text) < 0 Then Text3.Text = 0
    If Val(Text3.Text) > 256 Then Text3.Text = 256
    HScroll2.Value = Text3.Text
    
    Call HScroll2_Change
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    nKeyAscii = KeyAscii
    If KeyAscii = 13 Then
        Call Text3_Validate(True)
    End If
End Sub

Private Sub F1EN()
    Frame1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
End Sub

Private Sub F1DIS()
    Frame1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
End Sub

Private Sub F2EN()
    Frame2.Enabled = True
    Text1.Enabled = True
    HScroll1.Enabled = True
End Sub

Private Sub F2DIS()
    Frame2.Enabled = False
    Text1.Enabled = False
    HScroll1.Enabled = False
End Sub

Private Sub F3EN()
    Frame3.Enabled = True
    Label4.Enabled = True
    Text2.Enabled = True
    Combo1.Enabled = True
End Sub

Private Sub F3DIS()
    Frame3.Enabled = False
    Label4.Enabled = False
    Text2.Enabled = False
    Combo1.Enabled = False
End Sub

Private Sub F4EN()
    Frame4.Enabled = True
    Label6.Enabled = True
    Text3.Enabled = True
    HScroll2.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Label5.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    calculate.Enabled = True
    switch.Enabled = True
End Sub

Private Sub F4DIS()
    Frame4.Enabled = False
    Label6.Enabled = False
    Text3.Enabled = False
    HScroll2.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Label5.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    calculate.Enabled = False
    switch.Enabled = False
End Sub

Private Sub AllReset()
    For I = 1 To 25
        Label1(I).BackColor = &H80000013
    Next I
    
    Label2.Caption = "X = "
    Label3.Caption = "Y = "
    
    Text1.Text = 100
    HScroll1.Value = 100
    
    Text2.Text = 0
    Combo1.ListIndex = 0
    Call Combo1_Click
    
    Text3.Text = DefaultCut
    HScroll2.Value = DefaultCut
    Text4.Text = 0
    
    bSelecting = False
    cut = DefaultCut
    pic = 1
End Sub

Private Sub picReset()
    Picture1.Visible = True
    Picture3.Visible = False
    pic = 1
    Command2.Caption = "Switch to original image"
End Sub

