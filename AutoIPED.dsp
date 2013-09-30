# Microsoft Developer Studio Project File - Name="AutoIPED" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=AutoIPED - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "AutoIPED.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "AutoIPED.mak" CFG="AutoIPED - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "AutoIPED - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "AutoIPED - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/AutoIPED", ANQBAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "AutoIPED - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "D:\shareDLL\lib"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
F90=df.exe
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /ZI /Od /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 GetPropertyofMaterial.lib SelEngineVolume.lib FillExcelDll.lib UEwasp.lib CalcSupportSpan.lib CustomMaterial.lib /nologo /subsystem:windows /machine:I386
# SUBTRACT LINK32 /debug

!ELSEIF  "$(CFG)" == "AutoIPED - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
F90=df.exe
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 GetPropertyofMateriald.lib SelEngineVolumed.lib FillExcelDlld.lib UEwasp.lib CalcSupportSpand.lib CustomMateriald.lib /nologo /subsystem:windows /debug /machine:I386 /out:"D:\sharedll/AutoIPED.exe" /pdbtype:sept

!ENDIF 

# Begin Target

# Name "AutoIPED - Win32 Release"
# Name "AutoIPED - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AplicatioInitDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AutoIPED.cpp
# End Source File
# Begin Source File

SOURCE=.\AutoIPED.rc
# End Source File
# Begin Source File

SOURCE=.\AutoIPEDDoc.cpp
# End Source File
# Begin Source File

SOURCE=.\AutoIPEDView.cpp
# End Source File
# Begin Source File

SOURCE=.\browseforfolermodule.cpp
# End Source File
# Begin Source File

SOURCE=.\calc.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_Code.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_CodeCJJ.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_CodeGB.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_CodePC.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcAlphaOriginalData.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcOriginal.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThick_Economy.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThick_FaceTemp.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThick_HeatDensity.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThick_PreventCongeal.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThick_SubterraneanNoChannel.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThick_SubterraneanObstruct.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThickBase.cpp
# End Source File
# Begin Source File

SOURCE=.\CalcThickSubterranean.cpp
# End Source File
# Begin Source File

SOURCE=.\CalThread.cpp
# End Source File
# Begin Source File

SOURCE=.\column.cpp
# End Source File
# Begin Source File

SOURCE=.\columns.cpp
# End Source File
# Begin Source File

SOURCE=.\CoolControlsManager.cpp
# End Source File
# Begin Source File

SOURCE=.\dataformatdisp.cpp
# End Source File
# Begin Source File

SOURCE=.\datagrid.cpp
# End Source File
# Begin Source File

SOURCE=.\DataGridEx.cpp
# End Source File
# Begin Source File

SOURCE=.\DataShowDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\dialog_pform.cpp
# End Source File
# Begin Source File

SOURCE=.\DialogBarEx.cpp
# End Source File
# Begin Source File

SOURCE=.\DialogBarExContainer.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgAutoTotal.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgOption.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgPaintRuleModi.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgParameterCongeal.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgParameterSubterranean.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgSelCityWeather.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgSelCritDB.cpp
# End Source File
# Begin Source File

SOURCE=.\DlgShowAllTable.cpp
# End Source File
# Begin Source File

SOURCE=.\EDIBgbl.cpp
# End Source File
# Begin Source File

SOURCE=.\EditOilPaintDataDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\EditOriginalData.cpp
# End Source File
# Begin Source File

SOURCE=.\excel9.cpp
# End Source File
# Begin Source File

SOURCE=.\ExceptionInfoHandle.cpp
# End Source File
# Begin Source File

SOURCE=.\ExplainTable.cpp
# End Source File
# Begin Source File

SOURCE=.\font.cpp
# End Source File
# Begin Source File

SOURCE=.\FoxBase.cpp
# End Source File
# Begin Source File

SOURCE=.\FrmFolderLocation.cpp
# End Source File
# Begin Source File

SOURCE=.\HeatPreCal.cpp
# End Source File
# Begin Source File

SOURCE=.\HyperLink.cpp
# End Source File
# Begin Source File

SOURCE=.\ImportAutoPD.cpp
# End Source File
# Begin Source File

SOURCE=.\ImportDirFromXlsDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\ImportFromPreCalcXLSDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\ImportFromXLSDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\ImportPaintXLSDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Input.cpp
# End Source File
# Begin Source File

SOURCE=.\InputOtherData.cpp
# End Source File
# Begin Source File

SOURCE=.\MainFrm.cpp
# End Source File
# Begin Source File

SOURCE=.\MaterialName.cpp
# End Source File
# Begin Source File

SOURCE=.\mobject.cpp
# End Source File
# Begin Source File

SOURCE=.\MyButton.cpp
# End Source File
# Begin Source File

SOURCE=.\mylistctrl.cpp
# End Source File
# Begin Source File

SOURCE=.\newproject.cpp
# End Source File
# Begin Source File

SOURCE=.\OutExcelFileThread.cpp
# End Source File
# Begin Source File

SOURCE=.\Page1.cpp
# End Source File
# Begin Source File

SOURCE=.\Page2.cpp
# End Source File
# Begin Source File

SOURCE=.\ParseExpression.cpp
# End Source File
# Begin Source File

SOURCE=.\picture.cpp
# End Source File
# Begin Source File

SOURCE=.\ProgBar.cpp
# End Source File
# Begin Source File

SOURCE=.\ProjectOperate.cpp
# End Source File
# Begin Source File

SOURCE=.\PropertyWnd.cpp
# End Source File
# Begin Source File

SOURCE=.\RangeDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\ReportCalErrorDlgBar.cpp
# End Source File
# Begin Source File

SOURCE=.\ScrollTip.cpp
# End Source File
# Begin Source File

SOURCE=.\selbookmarks.cpp
# End Source File
# Begin Source File

SOURCE=.\split.cpp
# End Source File
# Begin Source File

SOURCE=.\splits.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\stddataformatsdisp.cpp
# End Source File
# Begin Source File

SOURCE=.\ToExcel.cpp
# End Source File
# Begin Source File

SOURCE=.\TotalTableIPED.cpp
# End Source File
# Begin Source File

SOURCE=.\VtoT.cpp
# End Source File
# Begin Source File

SOURCE=.\XDialog.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AplicatioInitDlg.h
# End Source File
# Begin Source File

SOURCE=.\AutoIPED.h
# End Source File
# Begin Source File

SOURCE=.\AutoIPEDDoc.h
# End Source File
# Begin Source File

SOURCE=.\AutoIPEDView.h
# End Source File
# Begin Source File

SOURCE=.\browseforfolermodule.h
# End Source File
# Begin Source File

SOURCE=.\calc.h
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_Code.h
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_CodeCJJ.h
# End Source File
# Begin Source File

SOURCE=.\CalcAlpha_CodePC.h
# End Source File
# Begin Source File

SOURCE=.\CalcAlphaOriginalData.h
# End Source File
# Begin Source File

SOURCE=.\CalcOriginal.h
# End Source File
# Begin Source File

SOURCE=.\CalcSupportSpanDll.h
# End Source File
# Begin Source File

SOURCE=.\CalcThick_Economy.h
# End Source File
# Begin Source File

SOURCE=.\CalcThick_FaceTemp.h
# End Source File
# Begin Source File

SOURCE=.\CalcThick_HeatDensity.h
# End Source File
# Begin Source File

SOURCE=.\CalcThick_PreventCongeal.h
# End Source File
# Begin Source File

SOURCE=.\CalcThick_SubterraneanNoChannel.h
# End Source File
# Begin Source File

SOURCE=.\CalcThick_SubterraneanObstruct.h
# End Source File
# Begin Source File

SOURCE=.\CalcThickBase.h
# End Source File
# Begin Source File

SOURCE=.\CalcThickSubterranean.h
# End Source File
# Begin Source File

SOURCE=.\CalThread.h
# End Source File
# Begin Source File

SOURCE=.\column.h
# End Source File
# Begin Source File

SOURCE=.\columns.h
# End Source File
# Begin Source File

SOURCE=.\CoolControlsManager.h
# End Source File
# Begin Source File

SOURCE=.\dataformatdisp.h
# End Source File
# Begin Source File

SOURCE=.\datagrid.h
# End Source File
# Begin Source File

SOURCE=.\DataGridEx.h
# End Source File
# Begin Source File

SOURCE=.\DataShowDlg.h
# End Source File
# Begin Source File

SOURCE=.\Dialog_pform.h
# End Source File
# Begin Source File

SOURCE=.\DialogBarEx.h
# End Source File
# Begin Source File

SOURCE=.\DialogBarExContainer.h
# End Source File
# Begin Source File

SOURCE=.\DlgAutoTotal.h
# End Source File
# Begin Source File

SOURCE=.\DlgOption.h
# End Source File
# Begin Source File

SOURCE=.\DlgPaintRuleModi.h
# End Source File
# Begin Source File

SOURCE=.\DlgParameterCongeal.h
# End Source File
# Begin Source File

SOURCE=.\DlgParameterSubterranean.h
# End Source File
# Begin Source File

SOURCE=.\DlgSelCityWeather.h
# End Source File
# Begin Source File

SOURCE=.\DlgSelCritDB.h
# End Source File
# Begin Source File

SOURCE=.\DlgShowAllTable.h
# End Source File
# Begin Source File

SOURCE=.\EDIBgbl.h
# End Source File
# Begin Source File

SOURCE=.\EditOilPaintDataDlg.h
# End Source File
# Begin Source File

SOURCE=.\EditOriginalData.h
# End Source File
# Begin Source File

SOURCE=.\excel9.h
# End Source File
# Begin Source File

SOURCE=.\ExceptionInfoHandle.h
# End Source File
# Begin Source File

SOURCE=.\ExplainTable.h
# End Source File
# Begin Source File

SOURCE=.\font.h
# End Source File
# Begin Source File

SOURCE=.\FoxBase.h
# End Source File
# Begin Source File

SOURCE=.\FrmFolderLocation.h
# End Source File
# Begin Source File

SOURCE=.\HeatPreCal.h
# End Source File
# Begin Source File

SOURCE=.\htmlhelp.h
# End Source File
# Begin Source File

SOURCE=.\HyperLink.h
# End Source File
# Begin Source File

SOURCE=.\ImportAutoPD.h
# End Source File
# Begin Source File

SOURCE=.\ImportDirFromXlsDlg.h
# End Source File
# Begin Source File

SOURCE=.\ImportFromPreCalcXLSDlg.h
# End Source File
# Begin Source File

SOURCE=.\ImportFromXLSDlg.h
# End Source File
# Begin Source File

SOURCE=.\ImportPaintXLSDlg.h
# End Source File
# Begin Source File

SOURCE=.\Input.h
# End Source File
# Begin Source File

SOURCE=.\InputOtherData.h
# End Source File
# Begin Source File

SOURCE=.\MainFrm.h
# End Source File
# Begin Source File

SOURCE=.\MaterialName.h
# End Source File
# Begin Source File

SOURCE=.\mobject.h
# End Source File
# Begin Source File

SOURCE=.\MyButton.h
# End Source File
# Begin Source File

SOURCE=.\mylistctrl.h
# End Source File
# Begin Source File

SOURCE=.\newproject.h
# End Source File
# Begin Source File

SOURCE=.\OutExcelFileThread.h
# End Source File
# Begin Source File

SOURCE=.\Page1.h
# End Source File
# Begin Source File

SOURCE=.\Page2.h
# End Source File
# Begin Source File

SOURCE=.\ParseExpression.h
# End Source File
# Begin Source File

SOURCE=.\picture.h
# End Source File
# Begin Source File

SOURCE=.\ProgBar.h
# End Source File
# Begin Source File

SOURCE=.\Project.h
# End Source File
# Begin Source File

SOURCE=.\ProjectOperate.h
# End Source File
# Begin Source File

SOURCE=.\PropertyWnd.h
# End Source File
# Begin Source File

SOURCE=.\RangeDlg.h
# End Source File
# Begin Source File

SOURCE=.\ReportCalErrorDlgBar.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\ScrollTip.h
# End Source File
# Begin Source File

SOURCE=.\selbookmarks.h
# End Source File
# Begin Source File

SOURCE=.\split.h
# End Source File
# Begin Source File

SOURCE=.\splits.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\stddataformatsdisp.h
# End Source File
# Begin Source File

SOURCE=.\ToExcel.h
# End Source File
# Begin Source File

SOURCE=.\TotalTableIPED.h
# End Source File
# Begin Source File

SOURCE=.\UEwasp.h
# End Source File
# Begin Source File

SOURCE=.\VtoT.h
# End Source File
# Begin Source File

SOURCE=.\XDialog.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\res\AutoIPED.ico
# End Source File
# Begin Source File

SOURCE=.\res\AutoIPED.rc2
# End Source File
# Begin Source File

SOURCE=.\res\AutoIPEDDoc.ico
# End Source File
# Begin Source File

SOURCE=.\res\bitmap10.bmp
# End Source File
# Begin Source File

SOURCE=.\res\bitmap8.bmp
# End Source File
# Begin Source File

SOURCE=.\res\DELETE.BMP
# End Source File
# Begin Source File

SOURCE=.\res\FRSREC_S.BMP
# End Source File
# Begin Source File

SOURCE=.\res\LSTREC_S.BMP
# End Source File
# Begin Source File

SOURCE=.\res\NEW.BMP
# End Source File
# Begin Source File

SOURCE=.\res\NXTREC_S.BMP
# End Source File
# Begin Source File

SOURCE=.\res\PRVREC_S.BMP
# End Source File
# Begin Source File

SOURCE=.\res\SAVE.BMP
# End Source File
# Begin Source File

SOURCE=.\res\splith.cur
# End Source File
# Begin Source File

SOURCE=.\res\splitv.cur
# End Source File
# Begin Source File

SOURCE=.\res\Toolbar.bmp
# End Source File
# Begin Source File

SOURCE=.\res\WZDELETE.BMP
# End Source File
# End Group
# Begin Group "DLL"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\UEwasp.dll
# End Source File
# End Group
# Begin Group "LIB"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\ODBCCP32.LIB
# End Source File
# Begin Source File

SOURCE=.\htmlhelp.lib
# End Source File
# Begin Source File

SOURCE=.\UEwasp.lib
# End Source File
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
# Section AutoIPED : {BEF6E003-A874-101A-8BBA-00AA00300CAB}
# 	2:5:Class:COleFont
# 	2:10:HeaderFile:font.h
# 	2:8:ImplFile:font.cpp
# End Section
# Section AutoIPED : {CDE57A52-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:5:Class:CSelBookmarks
# 	2:10:HeaderFile:selbookmarks.h
# 	2:8:ImplFile:selbookmarks.cpp
# End Section
# Section AutoIPED : {7BF80981-BF32-101A-8BBB-00AA00300CAB}
# 	2:5:Class:CPicture
# 	2:10:HeaderFile:picture.h
# 	2:8:ImplFile:picture.cpp
# End Section
# Section AutoIPED : {CDE57A41-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:5:Class:CDataGrid
# 	2:10:HeaderFile:datagrid.h
# 	2:8:ImplFile:datagrid.cpp
# End Section
# Section AutoIPED : {CDE57A50-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:5:Class:CColumns
# 	2:10:HeaderFile:columns.h
# 	2:8:ImplFile:columns.cpp
# End Section
# Section AutoIPED : {CDE57A54-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:5:Class:CSplit
# 	2:10:HeaderFile:split.h
# 	2:8:ImplFile:split.cpp
# End Section
# Section AutoIPED : {E675F3F0-91B5-11D0-9484-00A0C91110ED}
# 	2:5:Class:CDataFormatDisp
# 	2:10:HeaderFile:dataformatdisp.h
# 	2:8:ImplFile:dataformatdisp.cpp
# End Section
# Section AutoIPED : {99FF4676-FFC3-11D0-BD02-00C04FC2FB86}
# 	2:5:Class:CStdDataFormatsDisp
# 	2:10:HeaderFile:stddataformatsdisp.h
# 	2:8:ImplFile:stddataformatsdisp.cpp
# End Section
# Section AutoIPED : {CDE57A4F-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:5:Class:CColumn
# 	2:10:HeaderFile:column.h
# 	2:8:ImplFile:column.cpp
# End Section
# Section AutoIPED : {CDE57A53-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:5:Class:CSplits
# 	2:10:HeaderFile:splits.h
# 	2:8:ImplFile:splits.cpp
# End Section
# Section AutoIPED : {CDE57A43-8B86-11D0-B3C6-00A0C90AEA82}
# 	2:21:DefaultSinkHeaderFile:datagrid.h
# 	2:16:DefaultSinkClass:CDataGrid
# End Section
