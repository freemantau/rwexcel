HA$PBExportHeader$tplibtest.sra
$PBExportComments$Generated Application Object
forward
global type tplibtest from application
end type
global transaction sqlca
global dynamicdescriptionarea sqlda
global dynamicstagingarea sqlsa
global error error
global message message
end forward

global type tplibtest from application
string appname = "tplibtest"
end type
global tplibtest tplibtest

on tplibtest.create
appname="tplibtest"
message=create message
sqlca=create transaction
sqlda=create dynamicdescriptionarea
sqlsa=create dynamicstagingarea
error=create error
end on

on tplibtest.destroy
destroy(sqlca)
destroy(sqlda)
destroy(sqlsa)
destroy(error)
destroy(message)
end on

event open;open(w_main)
end event

