VERSION 5.00
Object = "{FFFD4965-0DF8-4B14-BBC9-CACEDCFC370F}#24.0#0"; "pTaskInfo.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin pAppInfo.usrAppInfo usrAppInfo1 
      Height          =   4200
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   7408
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
usrAppInfo1.TaskName = App.EXEName & ".exe"
End Sub
