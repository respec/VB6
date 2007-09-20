Attribute VB_Name = "modNFF"
Option Explicit

Global Project As nffProject

Sub Main()
  Set Project = New nffProject
  Project.LoadNFFdatabase "C:\vbExperimental\libNFF\NFF.mdb"
  Project.fileName = DefaultSaveFile
  Project.XML = WholeFileString(Project.fileName)
  frmStart.Show
End Sub

