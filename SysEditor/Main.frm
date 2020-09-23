VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6120
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":030A
   ScaleHeight     =   6885
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   495
      Left            =   4920
      MouseIcon       =   "Main.frx":8ADDC
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox StringCall 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   1965
   End
   Begin VB.TextBox StringName 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Text            =   "RegisteredOwner"
      Top             =   6960
      Width           =   1965
   End
   Begin VB.TextBox RegFolder 
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Text            =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
      Top             =   7320
      Width           =   1965
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   120
      MouseIcon       =   "Main.frx":8B0E6
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   4560
      MouseIcon       =   "Main.frx":8B3F0
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   3480
      MouseIcon       =   "Main.frx":8B6FA
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   2640
      MouseIcon       =   "Main.frx":8BA04
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image 
      Height          =   255
      Left            =   960
      MouseIcon       =   "Main.frx":8BD0E
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5400
      MouseIcon       =   "Main.frx":8C018
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5760
      MouseIcon       =   "Main.frx":8C322
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Kinzuacountry Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3120
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2400
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3600
      MouseIcon       =   "Main.frx":8C62C
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Type what name you would like to appear in your System Properties window below and simply click Apply!!"
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   3000
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    SetStringValue RegFolder, StringName, StringCall
      MsgBox ("Successfully changed your computer name!! Please double check by navigating to your Control Panel folder and double clicking the System Icon.")
      If DisplayErrorMsg = True Then
MsgBox ("CRAP!!! there was a problem Buddy")
      End If
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Image2_Click()
MsgBox ("This clears the text only .. click the [X] in the upper right corner of this dialog box to exit!!")
StringCall.Text = ""
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Image3_Click()
MsgBox ("This program was created by Brandon Perkins. You can view his other little applications at http://www.kinzuacountry.com/downloads.php... Thanks!"), vbExclamation, Test
End Sub

Private Sub Image_Click()
MsgBox ("This section is Disabled"), vbOKOnly
End Sub

Private Sub Image4_Click()
MsgBox ("This section is Disabled"), vbOKOnly
End Sub

Private Sub Image5_Click()
MsgBox ("This section is Disabled"), vbOKOnly
End Sub

Private Sub Image6_Click()
MsgBox ("This section is Disabled"), vbOKOnly
End Sub
