VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Format, Copy Disk & Buka File Program Tertentu"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Settings Tab"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Appearance Tab"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Screen Saver Tab"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Background Tab"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kotak Dialog Open With"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Disk To Disk Dialog"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Membuka kotak Format Dialog"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Membuka kotak Format Dialog
  Call Shell("rundll32.exe shell32.dll,SHFormatDrive")
End Sub

Private Sub Command2_Click()
  'Membuka kotak Copy Disk To Disk Dialog
  Call Shell("rundll32.exe diskcopy.dll, DiskCopyRunDll 0,0", 1)
  'Keterangan: 0 yg pertama = dari drive yang mana
  '            0 yg kedua   = ke drive yang mana
  '            0 = drive a: ; 1 = drive b:
End Sub

Private Sub Command3_Click()
  'Membuka dengan (user memilih dengan program yang
  'mana untuk membuka file yg dipilih)
  Shell ("rundll32.exe shell32.dll, OpenAs_RunDLL c:\autoexec.bat")
End Sub

Private Sub Command4_Click()
  'Menampilkan Properties, Background Tab
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,0", 1)
End Sub

Private Sub Command5_Click()
  'Menampilkan Properties, Screen Saver Tab
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,1", 1)
End Sub

Private Sub Command6_Click()
  'Menampilkan Properties, Appearance Tab
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,2", 1)
End Sub

Private Sub Command7_Click()
  'Menampilkan Properties, Settings Tab
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,3", 1)
End Sub



Private Sub Form_Load()

End Sub
