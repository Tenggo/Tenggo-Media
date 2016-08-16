VERSION 5.00
Begin VB.Form frm_Pub_PopUp 
   BorderStyle     =   0  'None
   ClientHeight    =   210
   ClientLeft      =   495
   ClientTop       =   1875
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu MnuFGMPInsertion 
      Caption         =   "MnuFGMPInsertion"
      Visible         =   0   'False
      Begin VB.Menu MnuFreeze 
         Caption         =   "Freeze Col"
      End
      Begin VB.Menu MnuUnFreeze 
         Caption         =   "UnFreeze Col"
      End
      Begin VB.Menu mnuHideCol 
         Caption         =   "Hide Col"
      End
      Begin VB.Menu mnuShowHidenCol 
         Caption         =   "Show Hiden Col"
      End
      Begin VB.Menu mdiv1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStartPeriode 
         Caption         =   "Mark as Start Periode"
      End
      Begin VB.Menu MnuEndPeriode 
         Caption         =   "Mark as End Periode"
      End
      Begin VB.Menu MnuClearPeriode 
         Caption         =   "Clear Periode"
      End
      Begin VB.Menu mdiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_set_objective 
         Caption         =   "Set Objective"
      End
      Begin VB.Menu mnu_view_objective 
         Caption         =   "View Objective"
      End
      Begin VB.Menu mnu_view_id 
         Caption         =   "View ID"
      End
      Begin VB.Menu mdiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_refresh_rate 
         Caption         =   "Refresh Rate"
      End
      Begin VB.Menu mnuRateInfo 
         Caption         =   "Rate Info"
      End
      Begin VB.Menu mdiv4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Approve 
         Caption         =   "Approve"
      End
      Begin VB.Menu mnu_unapprove 
         Caption         =   "UnApprove"
      End
      Begin VB.Menu mdiv5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Cancel_IB 
         Caption         =   "Cancel IB"
      End
      Begin VB.Menu mnu_Show_Related_IB 
         Caption         =   "Show Related IB"
      End
      Begin VB.Menu mdiv6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_update_actual_budget 
         Caption         =   "Update Actual Budget"
      End
      Begin VB.Menu mdiv7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_unlock 
         Caption         =   "UnLock Cell"
      End
   End
End
Attribute VB_Name = "frm_Pub_PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  Frm_MPInsertion
' Fungsi Submodul       :  Untuk Edit Insertion
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  16 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'*****************************************************************************
Option Explicit

Public Sub mnu_Approve_Click()

End Sub

Public Sub mnu_Cancel_IB_Click()

End Sub

Public Sub mnu_refresh_rate_Click()

End Sub

Public Sub Mnu_Set_Objective_Click()
    frm_MPSetObjective.Show 1
End Sub

Public Sub mnu_Show_Related_IB_Click()

End Sub

Public Sub mnu_unapprove_Click()

End Sub

Public Sub mnu_unlock_Click()

End Sub

Public Sub Mnu_update_actual_budget_Click()
    frm_MPActualBudgetUpdate.Show 1
End Sub

Public Sub mnu_view_id_Click()

End Sub

Public Sub mnu_view_objective_Click()

End Sub

Public Sub MnuClearPeriode_Click()

End Sub

Public Sub MnuEndPeriode_Click()

End Sub

Public Sub MnuFreeze_Click()

End Sub

Public Sub mnuHideCol_Click()

End Sub

Public Sub mnuRateInfo_Click()

End Sub

Public Sub MnuShowHidenCol_Click()
    Frm_MPInsertionHidenCol.Show 1
End Sub

Public Sub MnuStartPeriode_Click()

End Sub
