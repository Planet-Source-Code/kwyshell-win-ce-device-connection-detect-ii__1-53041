VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kwyshell .Net Device Connection Detector"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "frmMain"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmRegular 
      Interval        =   100
      Left            =   4170
      Top             =   1200
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "&Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2580
      TabIndex        =   1
      Top             =   945
      Width           =   1125
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1305
      TabIndex        =   0
      Top             =   945
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Kwyshell DccMan Conntection Detector VisualBasic Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   690
      TabIndex        =   2
      Top             =   150
      Width           =   3630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objDccManSink As CDccManSink

Private Sub cmdAbout_Click()

    MsgBox "Kwyshell eMail:kwyshell@yahoo.com.tw", vbOKOnly, "Device Connection Detector"

End Sub

Private Sub cmdSetting_Click()

    m_objDccManSink.ShowCommSettings

End Sub

Private Sub Form_Load()

    Set m_objDccManSink = New CDccManSink

End Sub

Private Sub tmRegular_Timer()

    MDccManSinkImpl.DccManSinkImpl_CallProc

End Sub
