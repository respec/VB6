Attribute VB_Name = "modMachineSpecific"
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants

'Full path of GenScn executable - WinKeyDriver.exe is expected in same place
Global Const MACHINE_EXENAME = "c:\data\vbExperimental\genscn\bin\GenScn.exe"

'Data file to load when GenScn starts - can be "" to start w/o loading anything
Global Const MACHINE_EXECMD = "" '"c:\data\shena\shena.sta"

'Full path of Status.exe (Status monitor)
Global Const MACHINE_EXESTATUS = "c:\data\vbExperimental\genscn\bin\Status.exe"

'WinHSPF uses the following constants, GenScn does not use them
Global Const MACHINE_HSPFMSGMDB = "c:\vbExperimental\hspfinfo\hspfmsg.mdb"
Global Const MACHINE_POLLUTANTLIST = "c:\vbexperimental\winhspf\poltnt_2.prn"
Global Const MACHINE_HSPFMSGWDM = "c:\data\hspfmsg.wdm"
Global Const MACHINE_STARTERPATH = "c:\data\starter"
Global Const MACHINE_EXEWINHSPF = "c:\vbExperimental\winhspf\bin\winhspf.exe"

Global Const MACHINE_EXEWDMUTIL = "c:\vbExperimental\WDMUtil\WDMUtil.exe"
Global Const MACHINE_CMDWDMUTIL = ""

