# Microsoft Visual C++ generated build script - Do not modify

PROJ = WUPBAS
DEBUG = 0
PROGTYPE = 1
CALLER = c:\data\c\wup\wbrun\wbrun.exe
ARGS = 
DLLS = 
D_RCDEFINES = /d_DEBUG 
R_RCDEFINES = /dNDEBUG 
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
FIRSTC = WUPBAS.C    
FIRSTCPP =             
RC = rc
CFLAGS_D_WDLL = /nologo /Gs /G2 /W4 /Zi /ALw /Od /D "_DEBUG" /D "DEBUGOUTPUT" /FR /GD /Fd"WUPBAS.PDB"
CFLAGS_R_WDLL = /nologo /Gs /G3 /Gy /W4 /ALw /Ox /D "NDEBUG" /FR /GD 
LFLAGS_D_WDLL = /NOLOGO /NOD /NOE /PACKC:61440 /ALIGN:16 /ONERROR:NOEXE /CO /MAP:FULL
LFLAGS_R_WDLL = /NOLOGO /NOD /NOE /PACKC:61440 /ALIGN:16 /ONERROR:NOEXE /MAP:FULL
LIBS_D_WDLL = oldnames libw ldllcew commdlg.lib olecli.lib olesvr.lib shell.lib toolhelp.lib 
LIBS_R_WDLL = oldnames libw ldllcew commdlg.lib olecli.lib olesvr.lib shell.lib toolhelp.lib 
RCFLAGS = /nologo 
RESFLAGS = /nologo /31 
RUNFLAGS = 
DEFFILE = WUPBAS.DEF
OBJS_EXT = 
LIBS_EXT = 
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
SBRS = WUPBAS.SBR \
		WBEXPR.SBR \
		WBFNDOS.SBR \
		WBFNSTR.SBR \
		WBFNWIN.SBR \
		WBFUNC.SBR \
		WBMEM.SBR \
		WBMISC.SBR \
		WBPARSER.SBR \
		WBSTMNT.SBR \
		WBSTRING.SBR \
		WBVAR.SBR \
		WBFNTIME.SBR \
		WBFILEIO.SBR \
		WBCCEMUL.SBR \
		WBSLIST.SBR \
		WBLIB.SBR \
		WBFNOTHR.SBR \
		WBFNFILE.SBR \
		WBDLLCAL.SBR \
		WBTYPE.SBR \
		WBDECLAR.SBR \
		WBEXTDLL.SBR


WUPBAS_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbfileio.h \
	c:\data\c\wup\wupbas\wbdllcal.h \
	c:\data\c\wup\wupbas\wupbas.rcv


WBEXPR_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbdllcal.h \
	c:\data\c\wup\wupbas\wbtype.h \
	c:\data\c\wup\wupbas\wbmem.h


WBFNDOS_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbslist.h \
	c:\data\c\wup\wupbas\wbmem.h \
	c:\data\c\wup\wupbas\wbfnwin.h \
	c:\data\c\wup\wupbas\wbfndos.h


WBFNSTR_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbfnstr.h


WBFNWIN_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbslist.h \
	c:\data\c\wup\wupbas\wbfnwin.h


WBFUNC_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfnstr.h \
	c:\data\c\wup\wupbas\wbfndos.h \
	c:\data\c\wup\wupbas\wbfnwin.h \
	c:\data\c\wup\wupbas\wbfntime.h \
	c:\data\c\wup\wupbas\wbfnfile.h \
	c:\data\c\wup\wupbas\wbfnothr.h \
	c:\data\c\wup\wupbas\wbdllcal.h


WBMEM_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h


WBMISC_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h


WBPARSER_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h


WBSTMNT_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbfileio.h \
	c:\data\c\wup\wupbas\wbdllcal.h \
	c:\data\c\wup\wupbas\wbtype.h \
	c:\data\c\wup\wupbas\wbmem.h


WBSTRING_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h


WBVAR_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbtype.h


WUPBAS_RCDEP = c:\data\c\wup\wupbas\wupbas.rcv


WBFNTIME_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbfnstr.h


WBFILEIO_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbfileio.h


WBCCEMUL_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h


WBSLIST_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbslist.h \
	c:\data\c\wup\wupbas\wbfileio.h


WBLIB_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h \
	c:\data\c\wup\wupbas\wbfndos.h \
	c:\data\c\wup\wupbas\wbdllcal.h


WBFNOTHR_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbfunc.h


WBFNFILE_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbmem.h \
	c:\data\c\wup\wupbas\wbfunc.h \
	c:\data\c\wup\wupbas\wbslist.h \
	c:\data\c\wup\wupbas\wbfileio.h \
	c:\data\c\wup\wupbas\wbfnfile.h \
	c:\data\c\wup\wupbas\wbfndos.h \
	c:\data\c\wup\wupbas\wbfnwin.h


WBDLLCAL_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbtype.h \
	c:\data\c\wup\wupbas\wbmem.h


WBTYPE_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h \
	c:\data\c\wup\wupbas\wbtype.h \
	c:\data\c\wup\wupbas\wbmem.h


WBDECLAR_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h


WBEXTDLL_DEP = c:\data\c\wup\wupbas\wupbas.h \
	c:\data\c\wup\wupbas\vbapisub.h


all:	$(PROJ).DLL $(PROJ).BSC

WUPBAS.OBJ:	WUPBAS.C $(WUPBAS_DEP)
	$(CC) $(CFLAGS) $(CCREATEPCHFLAG) /c WUPBAS.C

WBEXPR.OBJ:	WBEXPR.C $(WBEXPR_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBEXPR.C

WBFNDOS.OBJ:	WBFNDOS.C $(WBFNDOS_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFNDOS.C

WBFNSTR.OBJ:	WBFNSTR.C $(WBFNSTR_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFNSTR.C

WBFNWIN.OBJ:	WBFNWIN.C $(WBFNWIN_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFNWIN.C

WBFUNC.OBJ:	WBFUNC.C $(WBFUNC_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFUNC.C

WBMEM.OBJ:	WBMEM.C $(WBMEM_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBMEM.C

WBMISC.OBJ:	WBMISC.C $(WBMISC_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBMISC.C

WBPARSER.OBJ:	WBPARSER.C $(WBPARSER_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBPARSER.C

WBSTMNT.OBJ:	WBSTMNT.C $(WBSTMNT_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBSTMNT.C

WBSTRING.OBJ:	WBSTRING.C $(WBSTRING_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBSTRING.C

WBVAR.OBJ:	WBVAR.C $(WBVAR_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBVAR.C

WUPBAS.RES:	WUPBAS.RC $(WUPBAS_RCDEP)
	$(RC) $(RCFLAGS) $(RCDEFINES) -r WUPBAS.RC

WBFNTIME.OBJ:	WBFNTIME.C $(WBFNTIME_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFNTIME.C

WBFILEIO.OBJ:	WBFILEIO.C $(WBFILEIO_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFILEIO.C

WBCCEMUL.OBJ:	WBCCEMUL.C $(WBCCEMUL_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBCCEMUL.C

WBSLIST.OBJ:	WBSLIST.C $(WBSLIST_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBSLIST.C

WBLIB.OBJ:	WBLIB.C $(WBLIB_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBLIB.C

WBFNOTHR.OBJ:	WBFNOTHR.C $(WBFNOTHR_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFNOTHR.C

WBFNFILE.OBJ:	WBFNFILE.C $(WBFNFILE_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBFNFILE.C

WBDLLCAL.OBJ:	WBDLLCAL.C $(WBDLLCAL_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBDLLCAL.C

WBTYPE.OBJ:	WBTYPE.C $(WBTYPE_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBTYPE.C

WBDECLAR.OBJ:	WBDECLAR.C $(WBDECLAR_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBDECLAR.C

WBEXTDLL.OBJ:	WBEXTDLL.C $(WBEXTDLL_DEP)
	$(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WBEXTDLL.C


$(PROJ).DLL::	WUPBAS.RES

$(PROJ).DLL::	WUPBAS.OBJ WBEXPR.OBJ WBFNDOS.OBJ WBFNSTR.OBJ WBFNWIN.OBJ WBFUNC.OBJ \
	WBMEM.OBJ WBMISC.OBJ WBPARSER.OBJ WBSTMNT.OBJ WBSTRING.OBJ WBVAR.OBJ WBFNTIME.OBJ \
	WBFILEIO.OBJ WBCCEMUL.OBJ WBSLIST.OBJ WBLIB.OBJ WBFNOTHR.OBJ WBFNFILE.OBJ WBDLLCAL.OBJ \
	WBTYPE.OBJ WBDECLAR.OBJ WBEXTDLL.OBJ $(OBJS_EXT) $(DEFFILE)
	echo >NUL @<<$(PROJ).CRF
WUPBAS.OBJ +
WBEXPR.OBJ +
WBFNDOS.OBJ +
WBFNSTR.OBJ +
WBFNWIN.OBJ +
WBFUNC.OBJ +
WBMEM.OBJ +
WBMISC.OBJ +
WBPARSER.OBJ +
WBSTMNT.OBJ +
WBSTRING.OBJ +
WBVAR.OBJ +
WBFNTIME.OBJ +
WBFILEIO.OBJ +
WBCCEMUL.OBJ +
WBSLIST.OBJ +
WBLIB.OBJ +
WBFNOTHR.OBJ +
WBFNFILE.OBJ +
WBDLLCAL.OBJ +
WBTYPE.OBJ +
WBDECLAR.OBJ +
WBEXTDLL.OBJ +
$(OBJS_EXT)
$(PROJ).DLL
$(MAPFILE)
C:\MSVC\LIB\+
C:\MSVC\MFC\LIB\+
c:\ndk\win\+
$(LIBS)
$(DEFFILE);
<<
	link $(LFLAGS) @$(PROJ).CRF
	$(RC) $(RESFLAGS) WUPBAS.RES $@
	@copy $(PROJ).CRF MSVC.BND
	implib /nowep $(PROJ).LIB $(PROJ).DLL

$(PROJ).DLL::	WUPBAS.RES
	if not exist MSVC.BND 	$(RC) $(RESFLAGS) WUPBAS.RES $@

run: $(PROJ).DLL
	$(PROJ) $(RUNFLAGS)


$(PROJ).BSC: $(SBRS)
	bscmake @<<
/o$@ $(SBRS)
<<
