VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "BTRS"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20250
   Icon            =   "Admin.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Admin.frx":424A
   ScaleHeight     =   10635
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu usermenu 
      Caption         =   "User-Reg"
   End
   Begin VB.Menu addbusmenu 
      Caption         =   "Buses"
   End
   Begin VB.Menu pnrmodmenu 
      Caption         =   "PNR Modification"
   End
   Begin VB.Menu report1 
      Caption         =   "Tickets report"
   End
   Begin VB.Menu mnulog 
      Caption         =   "Logout"
   End
   Begin VB.Menu exitmenu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addbusmenu_Click()
Form10.Show
Unload Me
End Sub

Private Sub exitmenu_Click()
Unload Me
End Sub

Private Sub mnulog_Click()
Form2.Show
Unload Me
End Sub

Private Sub pnrmodmenu_Click()
Form11.Show
Unload Me
End Sub

Private Sub report1_Click()
Form12.Show
Unload Me
End Sub

Private Sub usermenu_Click()
Form4.Show
End Sub
