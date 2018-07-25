VERSION 5.00
Begin VB.Form frmSourceCodeManager 
   Caption         =   "Source Code Manager"
   ClientHeight    =   8955
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11475
   Icon            =   "frmSourceCodeManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu m_proj 
      Caption         =   "&Project"
      Begin VB.Menu m_proj_newproj 
         Caption         =   "&New Empty Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu m_proj_open 
         Caption         =   "&Open Existing Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu m_proj_add 
         Caption         =   "&Add To Project"
         Begin VB.Menu m_proj_add_source 
            Caption         =   "&Add Source File..."
            Shortcut        =   ^A
         End
         Begin VB.Menu m_proj_add_resource 
            Caption         =   "Add &Resource File..."
         End
         Begin VB.Menu m_proj_add_files 
            Caption         =   "Add &Files To Project Resource..."
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu m_proj_save 
         Caption         =   "&Save Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu m_proj_saveas 
         Caption         =   "Save Project &As..."
      End
      Begin VB.Menu m_proj_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_proj_make 
         Caption         =   "&Make Project Output..."
         Shortcut        =   ^M
      End
      Begin VB.Menu m_proj_process 
         Caption         =   "&Process..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu m_proj_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu m_proj_settings 
         Caption         =   "Project Se&ttings..."
         Shortcut        =   ^P
      End
      Begin VB.Menu m_proj_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu m_proj_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmSourceCodeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim p As SourceCodeManager

Friend Sub SetParent(Parent As SourceCodeManager)
    Set p = Parent
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not p Is Nothing Then
        Dim Cancel1 As Boolean
        Cancel1 = False
        Call p.ClosingForm(Cancel1)
        Cancel = IIf(Cancel1, True, False)
    End If
End Sub
