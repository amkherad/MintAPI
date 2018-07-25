VERSION 5.00
Begin VB.Form frmGlobalization 
   Caption         =   "MintIntellicense Globalization"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mFile_NewGlobalizationFile 
         Caption         =   "&New"
      End
      Begin VB.Menu mFile_OpenGlobalizationFile 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mFile_SaveGlobalizationFile 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mFile_SaveAsGlobalizationFile 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mFile_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mFile_FileProperties 
         Caption         =   "&File Properties..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mFile_LastSep 
         Caption         =   "-"
      End
      Begin VB.Menu mFile_Quit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mCultures 
      Caption         =   "&Cultures"
      Begin VB.Menu m_Cultures_NewCulture 
         Caption         =   "&New Culture"
      End
   End
   Begin VB.Menu mPreferences 
      Caption         =   "&Preferences"
   End
   Begin VB.Menu mAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmGlobalization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
