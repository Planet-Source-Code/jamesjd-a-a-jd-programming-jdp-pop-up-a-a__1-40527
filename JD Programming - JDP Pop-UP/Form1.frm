VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JD Programming - JDP Pop-UP"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Text            =   "http://www.jdprogramming.cjb.net"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton ClearPopUP 
      Caption         =   "Start Over"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "--> Your Text Here <--"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "--> Your Text Here <--"
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "--> Your Text Here <--"
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "--> Your Text Here <--"
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton SendPopUP 
      Caption         =   "Send"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents msn As MsgrObject
Attribute msn.VB_VarHelpID = -1
' This tells the program that's it's an msn object

Private Sub SendPopUP_Click()
msn.LocalState = MSTATE_INVISIBLE
' This makes you appear offline on msn
msn.Services.PrimaryService.FriendlyName = (Text4.Text)
' This changes your msn name
msn.LocalState = MSTATE_ONLINE
' This makes you appear online on msn
msn.LocalState = MSTATE_INVISIBLE
' This makes you appear offline on msn
msn.Services.PrimaryService.FriendlyName = (Text3.Text)
' This changes your msn name
msn.LocalState = MSTATE_ONLINE
' This makes you appear online on msn
msn.LocalState = MSTATE_INVISIBLE
' This makes you appear offline on msn
msn.Services.PrimaryService.FriendlyName = (Text2.Text)
' This changes your msn name
msn.LocalState = MSTATE_ONLINE
' This makes you appear online on msn
msn.LocalState = MSTATE_INVISIBLE
' This makes you appear offline on msn
msn.Services.PrimaryService.FriendlyName = (Text1.Text)
' This changes your msn name
msn.LocalState = MSTATE_ONLINE
' This makes you appear online on msn
msn.Services.PrimaryService.FriendlyName = (Text5.Text)
' This changes your msn name back to what it was
End Sub

Private Sub ClearPopUP_Click()
Text1.Text = ""
' This clears the text in TextBox1
Text2.Text = ""
' This clears the text in TextBox2
Text3.Text = ""
' This clears the text in TextBox3
Text4.Text = ""
' This clears the text in TextBox4
End Sub

Private Sub Form_Load()
Set msn = New MsgrObject
' This starts the msn object running
Text5.Text = msn.Services.PrimaryService.FriendlyName
' This shows your msn nick name when it loads
End Sub
