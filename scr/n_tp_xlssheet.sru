HA$PBExportHeader$n_tp_xlssheet.sru
forward
global type n_tp_xlssheet from nonvisualobject
end type
end forward

global type n_tp_xlssheet from nonvisualobject native "tplib.dll"
public function  string version()
public function  boolean writestr(long row,long col,string str,n_tp_xlsformat format,uint celltype)
public function  boolean setcol(long colfirst,long collast,double width)
public function  boolean setrow(long row,double height)
public subroutine  setpicture(long row, long col, long pictureId, double scale, long offset_x, long offset_y)
public function  boolean setmerge(long rowfirst,long rowlast,long colfirst,long collast)
public subroutine  setdisplaygridlines(boolean show)
public function  boolean writenum(long row,long col,double value,n_tp_xlsformat format)
public function  boolean writeblank(long row,long col,n_tp_xlsformat format)
public function  boolean writeformula(long row,long col,string formula,n_tp_xlsformat format)
public function  boolean writestr(long row,long col,string str,n_tp_xlsformat format)
public function  long firstrow()
public function  long lastrow()
public function  long firstcol()
public function  long lastcol()
public function  boolean insertrow(long rowfirst,long rowlast,boolean updatenameranges)
public function  boolean insertcol(long colfirst,long collast,boolean updatenameranges)
public subroutine  setname(string name)
public function  string getname()
public subroutine  setprintarea(long rowfirst,long rowlast,long colfirst,long collast)
public subroutine  setlandscape(boolean landscape)
public subroutine  setprintfit(long wpaeges,long hpages)
public subroutine  setpaper(long paper)
public subroutine  sethcenter(boolean hcenter)
public subroutine  setvcenter(boolean vcenter)
public subroutine  setmarginleft(double margin)
public subroutine  setmarginright(double margin)
public subroutine  setmargintop(double margin)
public subroutine  setmarginbottom(double margin)
public subroutine  split(long row,long col)
public function  boolean removerow(long rowfirst,long rowlast,boolean updateNamedRanges)
public function  boolean removecol(long colfirst,long collast,boolean updateNamedRanges)
public subroutine  setprintgridlines(boolean printline)
public function  boolean setfooter(string footer,double margin)
public function  double colwidth(long col)
public function  double rowheight(long row)
public function  int celltype(long row,long col)
public function  boolean isdate(long row,long col)
public function  boolean isformula(long row,long col)
public function  string readstr(long row,long col)
public function  double readnum(long row,long col)
public function  boolean readbool(long row,long col)
public function  string readformula(long row,long col)
public function  boolean writenum(long row,long col,double value)
public function  boolean writeblank(long row,long col)
public function  boolean writeformula(long row,long col,string formula)
public function  boolean writestr(long row,long col,string str)
end type
global n_tp_xlssheet n_tp_xlssheet

on n_tp_xlssheet.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_tp_xlssheet.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

