# tplib_powerbuilder
Direct reading and writing Excel files with Powerbuilder
# Examples
* writing data
<details>
>>> <summary> show code</summary>
	
```cpp
n_tp_excel book
book = Create n_tp_excel

book.createxls( cxls.TYPE_XLS)
//
int logoid
logoid = book.addpicture( "logo.png")
//
////fonts
n_tp_xlsfont textfont
textfont = book.addfont( )
textfont.setsize( 8)
textfont.setname( "Arial")

n_tp_xlsfont titlefont
titlefont = book.addfont( textfont )
titlefont.setsize( 38)
titlefont.setcolor( cxls.COLOR_GRAY25)

n_tp_xlsfont font12,font10
font12 = book.addfont( textfont )
font10 = book.addfont( textfont)
font12.setsize( 12)
font10.setsize( 10)

////format
n_tp_xlsformat textformat
textformat = book.addformat( )
textformat.setfont( textfont)
textformat.setalignh( cxls.ALIGNH_LEFT)

n_tp_xlsformat titleformat
titleformat = book.addformat( )
titleformat.setfont( titlefont)
titleformat.setalignh( cxls.ALIGNH_RIGHT)

n_tp_xlsformat companyformat
companyformat = book.addformat( )
companyformat.setfont( font12)

n_tp_xlsformat dateformat
dateformat = book.addformat(textformat)
dateformat.setnumformat( book.addcustomnumformat( "[$-409]mmmm\ d\,\ yyyy;@"))

n_tp_xlsformat phoneformat
phoneformat = book.addformat(textformat)
phoneformat.setnumformat( book.addcustomnumformat( "[<=9999999]###\-####;\(###\)\ ###\-####"))

n_tp_xlsformat borderformat
borderformat = book.addformat(textformat )
borderformat.setborder( cxls.BORDERSTYLE_THIN)
borderformat.setbordercolor( cxls.COLOR_GRAY25)
borderformat.setalignv( cxls.ALIGNV_CENTER)

n_tp_xlsformat percentformat
percentformat = book.addformat(borderformat )
percentformat.setnumformat( book.addcustomnumformat( "#%_)"))
percentformat.setalignh( cxls.ALIGNH_RIGHT)

n_tp_xlsformat textrightformat
textrightformat = book.addformat( textformat )
textRightFormat.setAlignH(cxls.ALIGNH_RIGHT);
textRightFormat.setAlignV(cxls.ALIGNV_CENTER);

n_tp_xlsformat thankformat
thankformat = book.addformat( )
thankFormat.setFont(font10);
thankFormat.setAlignH(cxls.ALIGNH_CENTER);

n_tp_xlsformat dollarformat
dollarformat = book.addformat( borderformat )
dollarformat.setnumformat( book.addcustomnumformat( "_($* # ##0.00_);_($* (# ##0.00);_($* -??_);_(@_)"))
//actions
n_tp_xlssheet sheet
sheet = book.addsheet( "Sales Receipt表")

//TODO
sheet.setdisplaygridlines(false)

sheet.setCol(1, 1, 36)
sheet.setCol(0, 0, 10)
sheet.setCol(2, 4, 11)

sheet.setRow(2, 47.25)
sheet.writeStr(2, 1, "Sales Receipt", titleFormat)
sheet.setMerge(2, 2, 1, 4)
sheet.setPicture(2, 1, logoId,1.0,0,0)


sheet.writeStr(4, 0, "Apricot Ltd.", companyFormat)
sheet.writeStr(4, 3, "Date:", textFormat)

//TODO
sheet.writeFormula(4, 4, "TODAY()", dateFormat)

sheet.writeStr(5, 3, "Receipt #:", textFormat)

//TODO
sheet.writeNum(5, 4, 652, textFormat)

sheet.writeStr(8, 0, "Sold to:", textFormat)
sheet.writeStr(8, 1, "John Smith", textFormat)
sheet.writeStr(9, 1, "Pineapple Ltd.", textFormat)
sheet.writeStr(10, 1, "123 Dreamland Street", textFormat)
sheet.writeStr(11, 1, "Moema, 52674", textFormat)

//TODO
sheet.writeNum(12, 1, 2659872055, phoneFormat)

sheet.writeStr(14, 0, "Item #", textFormat)
sheet.writeStr(14, 1, "Description", textFormat)
sheet.writeStr(14, 2, "Qty", textFormat)
sheet.writeStr(14, 3, "Unit Price", textFormat)
sheet.writeStr(14, 4, "Line Total", textFormat)

int row,col
string s
for row = 15 to 37
	sheet.setRow(row, 15)
	for col = 0 to 2
		//TODO
		sheet.writeBlank(row, col, borderFormat)
	next
	//TODO
	sheet.writeBlank(row, 3, dollarFormat)

	//TODO
	//s = sprintf('IF(C{1}>0;ABS(C{2}*D{3});"")',row + 1,row + 1,row + 1)
	s = 'IF(C' +string(row + 1)+ '>0;ABS(C' +string(row + 1)+ '*D' +string(row + 1)+ ');"")'
	sheet.writeFormula(row, 4, s, dollarFormat)
next 



sheet.writeStr(38, 3, "Subtotal ", textRightFormat)
sheet.writeStr(39, 3, "Sales Tax ", textRightFormat)
sheet.writeStr(40, 3, "Total ", textRightFormat)
sheet.writeFormula(38, 4, "SUM(E16:E38)", dollarFormat)
sheet.writeNum(39, 4, 0.2, percentFormat)
sheet.writeFormula(40, 4, "E39+E39*E40", dollarFormat)
sheet.setRow(38, 15)
sheet.setRow(39, 15)
sheet.setRow(40, 15)

sheet.writeStr(42, 0, "Thank you for your business!", thankFormat)
sheet.setMerge(42, 42, 0, 4)

// items

sheet.writeNum(15, 0, 45, borderFormat)
sheet.writeStr(15, 1, "Grapes", borderFormat)
sheet.writeNum(15, 2, 250, borderFormat)
sheet.writeNum(15, 3, 4.5, dollarFormat)

sheet.writeNum(16, 0, 12, borderFormat)
sheet.writeStr(16, 1, "Bananas", borderFormat)
sheet.writeNum(16, 2, 480, borderFormat)
sheet.writeNum(16, 3, 1.4, dollarFormat)

sheet.writeNum(17, 0, 19, borderFormat)
sheet.writeStr(17, 1, "Apples", borderFormat)
sheet.writeNum(17, 2, 180, borderFormat)
sheet.writeNum(17, 3, 2.8, dollarFormat)

book.save("receipt.xls",false)
//book->release();
book.close( )
destroy book

Messagebox('','complete!')
```

</details>

