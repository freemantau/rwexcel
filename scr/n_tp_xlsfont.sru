HA$PBExportHeader$n_tp_xlsfont.sru
forward
global type n_tp_xlsfont from nonvisualobject
end type
end forward

global type n_tp_xlsfont from nonvisualobject native "tplib.dll"
public function  string version()
public subroutine  setsize(int size)
public subroutine  setitalic(boolean italic)
public subroutine  setstrikeout(boolean strikeout)
public subroutine  setcolor(long color)
public subroutine  setbold(boolean bold)
public subroutine  setunderline(long underline)
public function  boolean setname(string name)
end type
global n_tp_xlsfont n_tp_xlsfont

on n_tp_xlsfont.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_tp_xlsfont.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

