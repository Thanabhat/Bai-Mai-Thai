VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "���������"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7680
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = "�Ը���ҹ" & Chr(13) & Chr(10) & _
                        "1. �Դ�ٻ�Ҿ ���͡ File > Open" & Chr(13) & Chr(10) & _
                        "2. ��Ѻ��Ҵ�ٻ�Ҿ �¡������͹ᶺ Change image size (%)" & Chr(13) & Chr(10) & _
                        "3. �ҡ��鹵ç�����������Ǩ�ԧŧ���ٻ�Ҿ �����������Ǩ�ԧ�˹������������" & Chr(13) & Chr(10) & _
                        "4. ����ö����¹˹��¡���Ѵ�� ���͡ Line option - Length" & Chr(13) & Chr(10) & _
                        "5. ��Ѻ��Ҵ��������բͧ�ٻ������Фӹǳ �¡������͹ᶺ Color depth" & Chr(13) & Chr(10) & _
                        "       -> ��ҵ�ͧ�������������շ����Ѵ���ŧ ����͹价ҧ����" & Chr(13) & Chr(10) & _
                        "       -> ��ҵ�ͧ�������������շ����Ѵ�٧��� ����͹价ҧ���" & Chr(13) & Chr(10) & _
                        "6. �Ҿ�鹷�������¡����� Calculate" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "How to use" & Chr(13) & Chr(10) & _
                        "1. Open the picture select " & Chr(34) & "File > Open" & Chr(34) & "." & Chr(13) & Chr(10) & _
                        "2. To change image size by using scroll bar at " & Chr(34) & "Change image size (%)" & Chr(34) & "." & Chr(13) & Chr(10) & _
                        "3. Draw the line that know the real length and input the real length (mm)." & Chr(13) & Chr(10) & _
                        "4. To change the unit by " & Chr(34) & "Line option - Length" & Chr(34) & "." & Chr(13) & Chr(10) & _
                        "5. To change the color depth for calculate by moving scroll bar." & Chr(13) & Chr(10) & _
                        "       -> To decrease the color depth, move bar to the left." & Chr(13) & Chr(10) & _
                        "       -> To increase the color depth, move bar to the right." & Chr(13) & Chr(10) & _
                        "6. Using Calculate button for measuring leaf area."
                        
End Sub
