HA$PBExportHeader$n_tp_xlsformat.sru
forward
global type n_tp_xlsformat from nonvisualobject
end type
end forward

global type n_tp_xlsformat from nonvisualobject native "tplib.dll"
public function  string version()
public subroutine  setalignh(uint alignh)
public subroutine  setalignv(uint alignv)
public function  boolean setfont(n_tp_xlsfont font)
public subroutine  setnumformat(long numformat)
public subroutine  setwrap(boolean wrap)
public function  boolean setrotation(int rotation)
public subroutine  setindent(int indent)
public subroutine  setshrinktofix(boolean shfix)
public subroutine  setborder(long borderstyle)
public subroutine  setbordercolor(long color)
public subroutine  setborderleft(long borderstyle)
public subroutine  setborderright(long borderstyle)
public subroutine  setbordertop(long borderstyle)
public subroutine  setborderbottom(long borderstyle)
public subroutine  setborderleftcolor(long color)
public subroutine  setborderrightcolor(long color)
public subroutine  setbordertopcolor(long color)
public subroutine  setborderbottomcolor(long color)
public function  n_tp_xlsfont getfont()
public subroutine  setfillpattern(long pattern)
public subroutine  setpatternforegroundcolor(long color)
end type
global n_tp_xlsformat n_tp_xlsformat

on n_tp_xlsformat.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_tp_xlsformat.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

