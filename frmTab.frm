VERSION 5.00
Begin VB.Form frmTab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "For Testing"
   ClientHeight    =   6105
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   360
         ScaleHeight     =   795
         ScaleWidth      =   1035
         TabIndex        =   2
         Tag             =   "6"
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   5280
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00000080&
      Caption         =   "Cancel"
      Height          =   555
      Left            =   5760
      TabIndex        =   10
      Tag             =   "10"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   555
      Left            =   5760
      TabIndex        =   9
      Tag             =   "9"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Tag             =   "0"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "frmTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

