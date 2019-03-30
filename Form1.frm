VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dizi Takip"
   ClientHeight    =   3585
   ClientLeft      =   6690
   ClientTop       =   3960
   ClientWidth     =   6720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6720
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   14
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   13
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":08CA
      Left            =   240
      List            =   "Form1.frx":08E3
      TabIndex        =   8
      Text            =   "Dizinizin Gününü Seçin"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÇIK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TEMÝZLE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KAYDET"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GÖSTER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "5. Takip Edilen Dizi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "4. Takip Edilen Dizi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "3. Takip Edilen Dizi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "2. Takip Edilen Dizi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "1. Takip Edilen Dizi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type topla
gn As String
ug As String
bgg As String
ig As String
dg As String
bg As String
ag As String
End Type
Dim veri As topla
Private Sub Command1_Click()
On Error GoTo hata
Select Case Combo1.Text
Case "PAZARTESÝ"
Combo1.Text = 1
Label6.Caption = "PAZARTESÝ"
Case "SALI"
Combo1.Text = 2
Label6.Caption = "SALI"
Case "ÇARÞAMBA"
Combo1.Text = 3
Label6.Caption = "ÇARÞAMBA"
Case "PERÞEMBE"
Combo1.Text = 4
Label6.Caption = "PERÞEMBE"
Case "CUMA"
Combo1.Text = 5
Label6.Caption = "CUMA"
Case "CUMARTESÝ"
Combo1.Text = 6
Label6.Caption = "CUMARTESÝ"
Case "PAZAR"
Combo1.Text = 7
Label6.Caption = "PAZAR"
End Select
Open "c:\diziay.dat" For Random As #1
Get #1, Combo1.Text, veri
Combo1.Text = veri.gn
Text1.Text = veri.ug
Text2.Text = veri.bgg
Text3.Text = veri.ig
Text4.Text = veri.dg
Text5.Text = veri.bg
Text6.Text = veri.ag
Close #1: Exit Sub
hata: MsgBox "Eyi Bak Bakam Yazdýklana!", vbExclamation, "Hata": Close #1
End Sub

Private Sub Command2_Click()
Select Case Combo1.Text
Case "PAZARTESÝ"
Combo1.Text = 1
Label6.Caption = "PAZARTESÝ"
Case "SALI"
Combo1.Text = 2
Label6.Caption = "SALI"
Case "ÇARÞAMBA"
Combo1.Text = 3
Label6.Caption = "ÇARÞAMBA"
Case "PERÞEMBE"
Combo1.Text = 4
Label6.Caption = "PERÞEMBE"
Case "CUMA"
Combo1.Text = 5
Label6.Caption = "CUMA"
Case "CUMARTESÝ"
Combo1.Text = 6
Label6.Caption = "CUMARTESÝ"
Case "PAZAR"
Combo1.Text = 7
Label6.Caption = "PAZAR"
End Select
Open "c:\diziay.dat" For Random As #1
On Error GoTo hata
veri.gn = Combo1.Text
veri.ug = Text1.Text
veri.bgg = Text2.Text
veri.ig = Text3.Text
veri.dg = Text4.Text
veri.bg = Text5.Text
veri.ag = Text6.Text
Put #1, veri.gn, veri
Close #1: Exit Sub
hata: MsgBox "Eyi Bak Bakam Yazdýklana!", vbExclamation, "Hata": Close #1
End Sub


Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command4_Click()
End
End Sub


