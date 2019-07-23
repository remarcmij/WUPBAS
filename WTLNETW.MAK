# Microsoft Visual C++ generated build script - Do not modify

PROJ = WTLNETW
DEBUG = 0
PROGTYPE = 1
CALLER = wtrun.exe c:\wupdev\cdromini.wtb
ARGS = 
DLLS = wupbas.dll
D_RCDEFINES = -d_DEBUG
R_RCDEFINES = -dNDEBUG
ORIGIN = MSVC
ORIGIN_VER = 1.00
PROJPATH = C:\DATA\C\WUP\WUPBAS\
USEMFC = 0
CC = cl
CPP = cl
CXX = cl
CCREATEPCHFLAG = 
CPPCREATEPCHFLAG = 
CUSEPCHFLAG = 
CPPUSEPCHFLAG = 
FIRSTC = WTLNETW.C   
FIRSTCPP =             
RC = rc
CFLAGS_D_WDLL = /nologo /G2 /W4 /Zi /ASw /Od /D "_DEBUG" /FR /GD /Fd"WTLNETW.PDB"
CFLAGS_R_WDLL = /nologo /Gs /G2 /Gy /W4 /ASw /Ox /D "NDEBUG" /FR /GD 
LFLAGS_D_WDLL = /NOLOGO /NOD /NOE /PACKC:61440 /ALIGN:16 /ONERROR:NOEXE /CO /MAP:FULL
LFLAGS_R_WDLL = /NOLOGO /NOD /NOE /PACKC:61440 /ALIGN:16 /ONERROR:NOEXE /MAP:FULL
LIBS_D_WDLL = oldnames libw sdllcew nwcalls commdlg.lib olecli.lib olesvr.lib shell.lib toolhelp.lib 
LIBS_R_WDLL = oldnames libw sdllcew nwcalls commdlg.lib olecli.lib olesvr.lib shell.lib toolhelp.lib 
RCFLAGS = /nologo
RESFLAGS = /nologo
RUNFLAGS = 
DEFFILE = WTLNETW.DEF
OBJS_EXT = 
LIBS_EXT = WUPBAS.LIB 
!if "$(DEBUG)" == "1"
CFLAGS = $(CFLAGS_D_WDLL)
LFLAGS = $(LFLAGS_D_WDLL)
LIBS = $(LIBS_D_WDLL)
MAPFILE = nul
RCDEFINES = $(D_RCDEFINES)
!else
CFLAGS = $(CFLAGS_R_WDLL)
LFLAGS = $(LFLAGS_R_WDLL)
LIBS = $(LIBS_R_WDLL)
MAPFILE = nul
RCDEFINES = $(R_RCDEFINES)
!endif
!if [if exist MSVC.BND del MSVC.BND]
!endif
SBRS = WTLNETW.SBR


WUPBAS_DEP = 

WTLNETW_DEP = c:\data\c\wup\wupbas\vbapisub.h \
	c:\ndk\include\nwcaldef.h \
	c:\ndk\include\nwserver.h \
	c:\ndk\include\nwerror.h \
	c:\ndk\include\nwdpath.h \
	c:\ndk\include\nwconnec.h \
	c:\ndk\include\nwbindry.h \
	c:\ndk\include\nwalias.h


all:	$(PROJ).DLL $(PROJ).BSC

WTLNETW.OBJ:	WTLNETW.C $(WTLNETW_DEP)
	$(CC) $(CFLAGS) $(CCREATEPCHFLAG) /c WTLNETW.C

WTLNETW.RES:	WTLNETW.RC $(WTLNETW_RCDEP)
	$(RC) $(RCFLAGS) $(RCDEFINES) -r WTLNETW.RC


$(PROJ).DLL::	WTLNETW.RES

$(PROJ).DLL::	WTLNETW.OBJ $(OBJS_EXT) $(DEFFILE)
	echo >NUL @<<$(PROJ).CRF
WTLNETW.OBJ +
$(OBJS_EXT)
$(PROJ).DLL
$(MAPFILE)
C:\MSVC\LIB\+
C:\MSVC\MFC\LIB\+
c:\ndk\win\+
WUPBAS.LIB+
$(LIBS)
$(DEFFILE);
<<
	link $(LFLAGS) @$(PROJ).CRF
	$(RC) $(RESFLAGS) WTLNETW.RES $@
	@copy $(PROJ).CRF MSVC.BND
	implib /nowep $(PROJ).LIB $(PROJ).DLL

$(PROJ).DLL::	WTLNETW.RES
	if not exist MSVC.BND 	$(RC) $(RESFLAGS) WTLNETW.RES $@

run: $(PROJ).DLL
	$(PROJ) $(RUNFLAGS)


$(PROJ).BSC: $(SBRS)
	bscmake @<<
/o$@ $(SBRS)
<<
