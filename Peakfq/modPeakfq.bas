Attribute VB_Name = "modPeakfq"
Option Explicit
Public PfqPrj As New pfqProject
Public DefPfqPrj As New pfqProject
Public gIPC As ATCoIPC


Sub Main()
  Dim ff As New ATCoFindFile

  ff.SetDialogProperties "Please locate the PKFQWin Batch Executable 'PKFQBat.EXE'", App.path & "\PKFQBat.exe"
  ff.SetRegistryInfo "PKFQWin", "files", "PKFQBat.exe"
  PfqPrj.PFQExeFileName = ff.GetName

  ff.SetDialogProperties "Please locate PKFQWin help file 'PeakFQ.chm'", App.path & "\PeakFQ.chm"
  ff.SetRegistryInfo "PKFQWin", "files", "PeakFQ.chm"
  App.HelpFile = ff.GetName
    
  Set gIPC = New ATCoIPC
  gIPC.SendMonitorMessage "(Caption PKFQWin Status)"
  
  frmPeakfq.Show

End Sub
