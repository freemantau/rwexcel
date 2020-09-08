HA$PBExportHeader$n_tp_excel.sru
forward
global type n_tp_excel from nonvisualobject
end type
end forward

global type n_tp_excel from nonvisualobject native "tplib.dll"
public function  string version()
public function  boolean load(string sfile)
public function  long close()
public function  long getsheetcount()
public function  string getsheetname(long index)
public function  long getrowcount(long sheetindex)
public function  long getcolumncount(long sheetindex)
public function  int getcelltype(long sheetindex,long row,long column)
public function  string readcellstring(long sheetindex,long row,long column)
public function  double readcellnumber(long sheetindex,long row,long column)
public function  boolean readcellboolean(long sheetindex,long row,long column)
public function  datetime readcelldatetime(long sheetindex,long row,long column)
public function  boolean createxls(int xlstype)
public function  n_tp_xlssheet addsheet(string name)
public function  boolean save(string filename,boolean usetempfile)
public function  n_tp_xlsfont addfont()
public function  n_tp_xlsformat addformat()
public function  long addpicture(string filename)
public function  long addcustomnumformat(string cformat)
public function  n_tp_xlsfont addfont(n_tp_xlsfont font)
public function  n_tp_xlsformat addformat(n_tp_xlsformat format)
public function  n_tp_xlssheet getsheet(long index)
public function  double datepack(int year,int month,int day,int hour,int minute,int second,int msec)
public function  string getlasterrmsg()
public function  long getrow(long sheetindex,long row,ref string subs[])
public function  boolean dateunpack(double value,ref long year,ref long month,ref long day,ref long hour,ref long minute,ref long second,ref long msecond)
end type
global n_tp_excel n_tp_excel

on n_tp_excel.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_tp_excel.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

