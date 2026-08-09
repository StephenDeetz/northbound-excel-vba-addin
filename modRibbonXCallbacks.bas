Attribute VB_Name = "modRibbonXCallbacks"
' Copyright (c) 2026 Northbound Group
' Contact: stephendeetz@northboundgroup.com
' SPDX-License-Identifier: MIT
Option Explicit

' ==================================
' Ribbon Callbacks => Entry Points
' ==================================
' "NBPub_" prefix avoids Excel resolving these onAction names to the
' identically-named callbacks in the internal add-in when both are loaded.

'Callback for btnPrettyPrintMain onAction
Public Sub NBPub_PrettyPrintCallback(control As IRibbonControl)
  ' "Pretty Print / Minify Toggle" => "ToggleSel"
  ToggleSel
End Sub

'Callback for sptPrettyPrintMenu1 onAction
Public Sub NBPub_MinifySelectionCallback(control As IRibbonControl)
  ' "Minify Selection" => "MinifySel"
  MinifySel
End Sub

'Callback for sptPrettyPrintMenu2 onAction
Public Sub NBPub_MinifySheetCallback(control As IRibbonControl)
  ' "Minify Sheet" => "MinifyActiveSht"
  MinifyActiveSht
End Sub

'Callback for sptPrettyPrintMenu3 onAction
Public Sub NBPub_MinifyWorkbookCallback(control As IRibbonControl)
  ' "Minify Workbook" => "MinifyActiveWbk"
  MinifyActiveWbk
End Sub

'Callback for sptPrettyPrintMenu4 onAction
Public Sub NBPub_PrettyPrintSelectionCallback(control As IRibbonControl)
  ' "Pretty Print Selection" => "PrettyPrintSel"
  PrettyPrintSel
End Sub

'Callback for sptPrettyPrintMenu5 onAction
Public Sub NBPub_PrettyPrintSheetCallback(control As IRibbonControl)
  ' "Pretty Print Sheet" => "PrettyPrintActiveSht"
  PrettyPrintActiveSht
End Sub

'Callback for sptPrettyPrintMenu6 onAction
Public Sub NBPub_PrettyPrintWorkbookCallback(control As IRibbonControl)
  ' "Pretty Print Workbook" => "PrettyPrintActiveWbk"
  PrettyPrintActiveWbk
End Sub

'Callback for mnuAboutHelp onAction
Public Sub NBPub_AboutHelpCallback(control As IRibbonControl)
  ThisWorkbook.FollowHyperlink "https://github.com/StephenDeetz/northbound-excel-vba-addin/blob/main/README.md"
End Sub

'Callback for mnuAboutSourceCode onAction
Public Sub NBPub_AboutSourceCodeCallback(control As IRibbonControl)
  ThisWorkbook.FollowHyperlink "https://github.com/StephenDeetz/northbound-excel-vba-addin"
End Sub

'Callback for mnuAboutWebsite onAction
Public Sub NBPub_AboutWebsiteCallback(control As IRibbonControl)
  ThisWorkbook.FollowHyperlink "https://northboundgroup.com/?utm_source=excel_addin&utm_medium=ribbon&utm_campaign=about_menu"
End Sub
