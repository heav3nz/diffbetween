<!--#include file="aspsopconnect.asp"-->
<% server.scripttimeout = 3600 %>
<%

	passedPO=request.querystring("po")

	Private Const xlPrintNoComments = -4142
	Private Const xlUnderlineStyleNone = -4142
	Private Const xlAutomatic = -4105
	Private Const xlDiagonalDown = 5
	Private Const xlDiagonalUp = 6
	Private Const xlEdgeLeft = 7
	Private Const xlEdgeTop = 8
	Private Const xlEdgeBottom = 9
	Private Const xlEdgeRight = 10
	Private Const xlInsideVertical = 11
	Private Const xlInsideHorizontal = 12
	Private Const xlLandscape = 2
	Private Const xlPageBreakPreview = 2
	Private Const xlGeneral = 1
	Private Const xlBottom = -4107
	Private Const xlTop = -4160
	Private Const xlDown = -4121
	Private Const xlSolid = 1

	Private Const xlNone = -4142
	Private Const xlCenter = -4108
	Private Const xlLeft  = -4131

	Private Const xlContinuous = 1
	Private Const xlDash = -4115
	Private Const xlDashDot = 4
	Private Const xlDashDotDot = 5
	Private Const xlDot = -4118
	Private Const xlDouble = -4119
	Private Const xlLineStyleNone = -4142
	Private Const xlSlantDashDot = 13

	Private Const xlThin = 2
	Private Const xlThick = 4
	Private Const xlMedium = -4142
	Private Const xlHairline = 1

	Private Const xlDoNotSaveChanges = 2
	Private Const xlSaveChanges = 1
	dim variable
	dim discval
	dim pototval

	pay_mess = ""

	qry="select * from orderheader where po='" & passedPO  &"'"
	response.write(curclient)
	Set oRS = oConn.Execute(qry)

	curClient=trim(oRs.Fields("client"))
	curCountry=oRs.Fields("country")
	curLot = trim(oRs.Fields("lot"))
	curRequest = oRs.Fields("q_request")
	curDiscount = oRs.Fields("discount")
	curadjust = oRs.Fields("adjust")
	curAdjustDesc = oRs.Fields("adj_desc")
	curFreight = oRs.Fields("freight")
	curpricetype = trim(oRs.Fields("pricetype"))
	shipto = trim(oRs.Fields("shipto"))

	qry1="select po_mess,notes,contacto,telefono,direccion from clients where UPPER(rtrim(client))= '" & TRIM(ucase(curClient)) &"'"

	'qry1="select po_mess,notes,contacto,telefono,direccion from clients where UPPER(rtrim(client))like '%LLC HOSPITAL INTENSIVE%'"

	Set oRS1 = oConn.Execute(qry1)

	'response.write(curclient)
	'response.write(qry1)

	if not oRS1.eof then

		contacto = oRs1.Fields("contacto").value
		telefono = oRs1.Fields("telefono").value
		direccion = oRs1.Fields("direccion").value

	'response.write(curclient)
	'response.write(qry1)
	'response.write(contacto)

	else
	response.write("VACIO")

	end if
	'if curpricetype="E" then
		'SetLocale("es-es")
	'end if

	set oRS = Nothing
	set oRS1 = Nothing


'start


Private Const adSearchForward = 1
gethistory="Y"
qry="select * from orderheader where rtrim(client)="
qry = qry + "'" + trim(curClient) + "' "
qry = qry + " and "
qry = qry + "rtrim(country) = "
qry = qry + "'"+trim(curCountry)+"'"

Set TrackoRS = oConn.Execute(qry)
previoustr=""
thisPO=passedPO
if not TrackoRS.eof then


	doloop="Y"

	do while doloop = "Y"

		TrackoRS.MoveFirst
		'oRS.Find "547", 0, adSearchForward, 0

		'TrackoRS.Find "po='"+trim(thisPO)+"'", 0, adSearchForward, 1
		TrackoRS.Find "po like '%"+thisPO+"%'", 0, adSearchForward, 1
		if not TrackoRS.eof then
			thisPO=TrackoRS.Fields("transfered")

			if not trim(thispo) = "" then
				if trim(previoustr)="" then
					previoustr= " " & trim(thisPO) &  " "
				else
					previoustr = " " & trim(thisPO) &  " "  +  previoustr
				end if
			else

				doloop="N"
			end if
		else
			doloop= "N"
		end if
	loop
else

end if




thisPO=passedPO
afterstr = ""
'TrackoRS.MoveFirst
if not TrackoRS.eof then


	doloop="Y"

	do while doloop = "Y"

		TrackoRS.MoveFirst
		'oRS.Find "547", 0, adSearchForward, 0

		'TrackoRS.Find "transfered='"+trim(thisPO)+"'", 0, adSearchForward, 1
		TrackoRS.Find "transfered like '%"+thisPO+"%'", 0, adSearchForward, 1
		if not TrackoRS.eof then
			thisPO=TrackoRS.Fields("po")

			if not trim(thispo) = "" then
				if trim(afterstr)="" then
					afterstr = " " & trim(thisPO) & " "
				else
					afterstr =  afterstr & "  " & trim(thisPO)
				end if
			else

				doloop="N"
			end if
		else
			doloop= "N"
		end if
	loop
else

end if
trackStr= "Order History: " +previoustr +  " " + trim(passedPO) + " " + afterstr

'response.Write("It is in excelproforma")
'response.End()

'end


	set oExcel = Server.CreateObject("Excel.Application")
	set workbooks = oExcel.Workbooks
	set workbook  = workbooks.Add
	set Sheet = workbook.ActiveSheet





	'xls_dir = Server.MapPath("exceldump")
	xls_dir = "c:\Inetpub\wwwroot\sop\exceldump"
	'xls_dir = "C:\Inetpub\wwwroot\sop2\exceldump"
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if left(passedPO,1)="Q" then
		tempfile = fs.GetBaseName(fs.GetTempName()) & "_Quote " & trim(passedPO) & ".xls"
	else
		tempfile = fs.GetBaseName(fs.GetTempName()) & "_PO " & trim(passedPO) & ".xls"
	end if
	'tempfile=trim(passedPO)+".xls"
	xls_path = xls_dir & "\" & tempfile

	OutputType  = request.querystring("type")

	disp_mess = ""

	'if OutputType = "factory" then
		'qry="select fac_mess from clients where client='" & trim(curClient)  &"'"

		'Set oRS = oConn.Execute(qry)
		'disp_mess=trim(oRs.Fields("fac_mess"))

	'end if




	if OutputType = "proforma" then
		qry="select po_mess,notes from clients where upper(ltrim(rtrim(client)))='" & ucase(trim(curClient))  &"'"

		xxx=qry
		Set oRS = oConn.Execute(qry)
		if not oRS.eof then
			disp_mess=trim(oRs.Fields("po_mess"))
			cNotes = oRs.Fields("notes").value
			cNotes = ucase(cNotes)

			'response.write(qry)
			if instr(1,cNotes,"60 DAYS") > 0 then

				if ucase(trim(curCountry))= "MONTENEGRO" or  ucase(trim(curCountry))= "BULGARIA" and ucase(trim(curClient)) = "AMGPHARM" then
				'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "GREECE" and ucase(trim(curClient)) = "BIOCORE" then
					'pay_mess ="*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "TURKEY" and ucase(trim(curClient)) = "BEYBI" then
					'pay_mess = "*Payment 30 Days from Invoice"


				elseif ucase(trim(curCountry))= "HONDURAS" or ucase(trim(curCountry))= "EL SALVADOR" or ucase(trim(curCountry))= "SALVADOR" or ucase(trim(curCountry))= "GUATEMALA" or ucase(trim(curCountry))= "NICARAGUA" or ucase(trim(curCountry))= "COSTA RICA" or ucase(trim(curCountry))= "PANAMA" and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTRO AMERICA S A" or ucase(trim(curClient)) = "HEALTHCARE PRODUCTS" then
					'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "EGYPT" or  ucase(trim(curCountry))= "SAUDI ARABIA" and ucase(trim(curClient)) = "IM TRADING & MARKETING" then
				'pay_mess = "*Payment 60 Days from Invoice"


				elseif ucase(trim(curCountry))= "EGYPT" or  ucase(trim(curCountry))= "SAUDI ARABIA" and ucase(trim(curClient)) = "IM TRADING & MARKETING" then
				'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "RUSSIA" and ucase(trim(curClient)) = "LLC HOSPITAL INTENSIVE" then
					'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "SRI LANKA" and ucase(trim(curClient)) = "MARKSS HLC PVT LTD" then
					'pay_mess ="*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "VIETNAM" and ucase(trim(curClient)) = "MINH PHUONG" then
					'pay_mess ="*Payment 60 Days from Invoice"


				elseif ucase(trim(curCountry))= "NETTHERLANDS" or ucase(trim(curCountry))= "THE NETTHERLANDS" and ucase(trim(curClient)) = "QRSHC" then
					'pay_mess = "*Payment 30 Days from Invoice"

				elseif ucase(trim(curCountry))= "POLAND" and ucase(trim(curClient)) = "TOP MED" then
					'pay_mess ="*Payment 90 Days from Invoice"


				else
			   pay_mess = "*Payment 60 Days from Invoice"
			   end if
			else
				if ucase(trim(curCountry))= "MONTENEGRO" or  ucase(trim(curCountry))= "BULGARIA" and ucase(trim(curClient)) = "AMGPHARM" then
				'pay_mess = "*Payment 60 Days from Invoice"


				elseif ucase(trim(curCountry))= "GREECE" and ucase(trim(curClient)) = "BIOCORE" then
					'pay_mess ="*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "TURKEY" and ucase(trim(curClient)) = "BEYBI" then
					'pay_mess = "*Payment 30 Days from Invoice"

				elseif ucase(trim(curCountry))= "HONDURAS" or ucase(trim(curCountry))= "EL SALVADOR" or ucase(trim(curCountry))= "SALVADOR" or ucase(trim(curCountry))= "GUATEMALA" or ucase(trim(curCountry))= "NICARAGUA" or ucase(trim(curCountry))= "COSTA RICA" or ucase(trim(curCountry))= "PANAMA" and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTRO AMERICA S A" or ucase(trim(curClient)) = "HEALTHCARE PRODUCTS" then
					'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "EGYPT" or  ucase(trim(curCountry))= "SAUDI ARABIA" and ucase(trim(curClient)) = "IM TRADING & MARKETING" then
				'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "RUSSIA" and ucase(trim(curClient)) = "LLC HOSPITAL INTENSIVE" then
					'pay_mess = "*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "SRI LANKA" and ucase(trim(curClient)) = "MARKSS HLC PVT LTD" then
					'pay_mess ="*Payment 60 Days from Invoice"


				elseif ucase(trim(curCountry))= "VIETNAM" and ucase(trim(curClient)) = "MINH PHUONG" then
					'pay_mess ="*Payment 60 Days from Invoice"

				elseif ucase(trim(curCountry))= "NETHERLANDS" or ucase(trim(curCountry))= "THE NETHERLANDS" and ucase(trim(curClient)) = "QRSHC" then
					'pay_mess = "*Payment 30 Days from Invoice"

				elseif ucase(trim(curCountry))= "POLAND" and ucase(trim(curClient)) = "TOP MED" then
					'pay_mess ="*Payment 60 Days from Invoice"

				else
				pay_mess = ""
				end if
			end if
		else
			disp_mess=""
		end if

	end if



	donoteheader="Y"

	IF  OutputType="invoice" then
		datastart=23
	else
		datastart=23
	end if

		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\logo3.jpg").Select

	if not OutputType = "proforma" then
		oExcel.Selection.ShapeRange.ScaleWidth 0.58, msoFalse, msoScaleFromTopLeft
	    oExcel.Selection.ShapeRange.ScaleHeight 0.58, msoFalse, msoScaleFromTopLeft
	end if

	oExcel.cells(6,1).VALUE = "14175 NW 60th Ave"
	oExcel.cells(7,1).VALUE = "Miami Lakes, FL 33014"
	oExcel.cells(8,1).VALUE = "Tel: 305-824-1048"
	oExcel.cells(9,1).VALUE =  "Fax: 305-437-7607"


	'oExcel.cells(2,10).VALUE = "LuisJr@demetech.us"
	oExcel.ActiveSheet.Hyperlinks.Add oExcel.Range("H2"), "mailto:luisjr@demetech.us", , ,"LuisJr@demetech.us"



	'oExcel.cells(3,10).VALUE = "www.demetech.us"
	oExcel.ActiveSheet.Hyperlinks.Add oExcel.Range("H3"), "www.demetech.us", , ,"www.demetech.us"

	oExcel.ActiveSheet.Range("H2").font.bold=True
	oExcel.ActiveSheet.Range("H3").font.bold=True





	Select Case OutputType


		Case "factory":


			oExcel.cells(17,1).VALUE = "Costs in US Dollars and for box of 12 units."
			oExcel.ActiveSheet.Range("Q1").font.bold=True


		Case "factorycomp2":


			oExcel.cells(17,1).VALUE = "Costs in US Dollars and for box of 12 units."
			oExcel.ActiveSheet.Range("Q1").font.bold=True


		Case "factorycomp":


			oExcel.cells(17,1).VALUE = "Costs in US Dollars and for box of 12 units."
			oExcel.ActiveSheet.Range("Q1").font.bold=True


		Case "proforma":

			if ucase(trim(curCountry)) = "HONDURAS" and ucase(trim(shipto)) = "PANAMA" and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTROAMERICA, S. DE R.L." then
																													'Healthcare Products Centroamerica, S. de R.L.

			oExcel.cells(11,1).VALUE = "Quoted to: " & trim(curRequest) & "  " & trim(curClient)
			'oExcel.cells(12,11).VALUE = "For Sale and Use in " & trim(curCountry)
			oExcel.cells(21,1).VALUE = "This is our price confirmation. "
			oExcel.cells(12,1).VALUE = "Jimmy Zonta                      "
			oExcel.cells(13,1).VALUE = "Colonia Roble Oeste,                            "
			oExcel.cells(14,1).VALUE = "1 era Entrada a Mano Izquierda                              "
			oExcel.cells(15,1).VALUE = "Tegucigalpa"
			oExcel.cells(16,1).VALUE = "HONDURAS"
			oExcel.cells(17,1).VALUE = "For Sale and Use in PANAMA" '& trim(curCountry)
			'oExcel.cells(17,1).VALUE = "(Approved ship to Honduras, El Salvador,"
			'oExcel.cells(18,1).VALUE = "Guatemala and Panama)"

			oExcel.cells(13,1).font.bold = true
			oExcel.cells(12,11).font.bold = true
			oExcel.cells(14,1).font.bold = true
			oExcel.cells(15,1).font.bold = true
			oExcel.cells(16,1).font.bold = true
			oExcel.cells(17,1).font.bold = true
			oExcel.cells(18,1).font.bold = true
			oExcel.cells(21,1).font.bold = true

			'oExcel.cells(11,12).VALUE = "Ship to"
			'oExcel.cells(12,12).VALUE =  trim(curClient)
			'oExcel.cells(13,12).VALUE = "Lic. Juan Jos� Zontas Sing "
			'oExcel.cells(14,12).VALUE = "Calle 122 Via Jos� Agust�n Arango"
			'oExcel.cells(15,12).VALUE = "Corregimiento de Juan D�az, "
			'oExcel.cells(16,12).VALUE = "Centro Industrial Balzac, "
			'oExcel.cells(17,12).VALUE = "Local No. 4 Ciudad de Panam�, Panam� "
			'oExcel.cells(13,12).font.bold = true
			'oExcel.cells(11,12).font.bold = true
			'oExcel.cells(15,12).font.bold = true
			'oExcel.cells(16,12).font.bold = true
			'oExcel.cells(12,12).font.bold = true
			'oExcel.cells(18,12).font.bold = true
			'oExcel.cells(18,12).font.bold = true

			elseif ucase(trim(curCountry)) = "HONDURAS" and ucase(trim(shipto)) = "HONDURAS" and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTROAMERICA, S. DE R.L." then

			oExcel.cells(11,1).VALUE = "Quoted to: " & trim(curRequest) & "  " & trim(curClient)
			'oExcel.cells(12,11).VALUE = "For Sale and Use in " & trim(curCountry)
			oExcel.cells(21,1).VALUE = "This is our price confirmation. "
			oExcel.cells(12,1).VALUE = "Jimmy Zonta                      "
			oExcel.cells(13,1).VALUE = "Colonia Roble Oeste,                                 "
			oExcel.cells(14,1).VALUE = "1 era Entrada a Mano Izquierda                              "
			oExcel.cells(15,1).VALUE = "Tegucigalpa"
			oExcel.cells(16,1).VALUE = "HONDURAS"
			oExcel.cells(17,1).VALUE = "For Sale and Use in HONDURAS"' & trim(curCountry)
			'oExcel.cells(17,1).VALUE = "(Approved ship to Honduras, El Salvador,"
			'oExcel.cells(18,1).VALUE = "Guatemala and Panama)"
			oExcel.cells(13,1).font.bold = true
			oExcel.cells(12,11).font.bold = true
			oExcel.cells(14,1).font.bold = true
			oExcel.cells(15,1).font.bold = true
			oExcel.cells(16,1).font.bold = true
			oExcel.cells(17,1).font.bold = true
			oExcel.cells(18,1).font.bold = true
			oExcel.cells(21,1).font.bold = true




			elseif ucase(trim(curCountry)) = "HONDURAS" and (ucase(trim(shipto)) = "EL SALVADOR" or ucase(trim(shipto)) = "SALVADOR") and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTROAMERICA, S. DE R.L." then

			oExcel.cells(11,1).VALUE = "Quoted to: " & trim(curRequest) & "  " & trim(curClient)
			'oExcel.cells(12,11).VALUE = "For Sale and Use in " & trim(curCountry)
			oExcel.cells(21,1).VALUE = "This is our price confirmation. "
			oExcel.cells(12,1).VALUE = "Jimmy Zonta                      "
			oExcel.cells(13,1).VALUE = "Colonia Roble Oeste,                                           "
			oExcel.cells(14,1).VALUE = "1 era Entrada a Mano Izquierda                                           "
			oExcel.cells(15,1).VALUE = "Tegucigalpa"
			oExcel.cells(16,1).VALUE = "HONDURAS"
			oExcel.cells(17,1).VALUE = "For Sale and Use in EL SALVADOR"' & trim(curCountry)
			'oExcel.cells(17,1).VALUE = "(Approved ship to Honduras, El Salvador,"
			'oExcel.cells(18,1).VALUE = "Guatemala and Panama)"
			oExcel.cells(13,1).font.bold = true
			oExcel.cells(12,11).font.bold = true
			oExcel.cells(14,1).font.bold = true
			oExcel.cells(15,1).font.bold = true
			oExcel.cells(16,1).font.bold = true
			oExcel.cells(17,1).font.bold = true
			oExcel.cells(18,1).font.bold = true
			oExcel.cells(21,1).font.bold = true



			elseif ucase(trim(curCountry)) = "HONDURAS" and (ucase(trim(shipto)) = "GUATEMALA") and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTROAMERICA, S. DE R.L." then

			oExcel.cells(11,1).VALUE = "Quoted to: " & trim(curRequest) & "  " & trim(curClient)
			'oExcel.cells(12,11).VALUE = "For Sale and Use in " & trim(curCountry)
			oExcel.cells(21,1).VALUE = "This is our price confirmation."
			oExcel.cells(12,1).VALUE = "Jimmy Zonta                      "
			oExcel.cells(13,1).VALUE = "Colonia Roble Oeste,                                                 "
			oExcel.cells(14,1).VALUE = "1 era Entrada a Mano Izquierda                                               "
			oExcel.cells(15,1).VALUE = "Tegucigalpa"
			oExcel.cells(16,1).VALUE = "HONDURAS"
			oExcel.cells(17,1).VALUE = "For Sale and Use in GUATEMALA"' & trim(curCountry)
			'oExcel.cells(17,1).VALUE = "(Approved ship to Honduras, El Salvador,"
			'oExcel.cells(18,1).VALUE = "Guatemala and Panama)"

			oExcel.cells(13,1).font.bold = true
			oExcel.cells(12,11).font.bold = true
			oExcel.cells(14,1).font.bold = true
			oExcel.cells(15,1).font.bold = true
			oExcel.cells(16,1).font.bold = true
			oExcel.cells(17,1).font.bold = true
			oExcel.cells(18,1).font.bold = true
			oExcel.cells(21,1).font.bold = true




			else
			if curClient="GENCASA" then
				curClient = "Genericos Del Caribe, S.R. L"

				end if
			oExcel.cells(11,1).VALUE = "Quoted to: " & trim(curRequest) & "  " & trim(curClient)
			oExcel.cells(18,1).VALUE = "For Sale and Use in " & trim(curCountry)
			if trim(curpricetype) ="E" then
				oExcel.cells(20,1).VALUE = "This is our price confirmation, prices in Euros and for box of 12 units."
				'oExcel.cells(22,1).VALUE = "Prices in Euros and for box of 12 units."
				'oExcel.ActiveSheet.Range("A22").HorizontalAlignment = xlLeft
			else
				oExcel.cells(20,1).VALUE = "This is our price confirmation, prices in US Dollars and for box of 12 units."
			end if

			oExcel.cells(12,1).VALUE = contacto
			'oExcel.cells(14,1).VALUE = telefono
			'oExcel.cells(13,1).VALUE = direccion
			oExcel.cells(13,1).font.bold = true
			oExcel.cells(14,1).font.bold = true
			oExcel.cells(17,1).font.bold = true
			oExcel.cells(16,1).font.bold = true
			oExcel.cells(19,1).font.bold = true
			oExcel.cells(20,1).font.bold = true
			oExcel.cells(21,1).font.bold = true
			oExcel.cells(22,1).font.bold = true


			end if

			' if trim(curpricetype) ="E" then

				' oExcel.cells(22,1).VALUE = "Prices in Euros and for box of 12 units."
				' oExcel.ActiveSheet.Range("A22").HorizontalAlignment = xlLeft
			' end if
			'else
				'oExcel.cells(22,1).VALUE = "Prices in US Dollars and for box of 12 units."
				'oExcel.cells(22,1).HorizontalAlignment = xlLeft
				'oExcel.ActiveSheet.Range("H2").HorizontalAlignment = xlCenter
				'oExcel.ActiveSheet.Range("A22").HorizontalAlignment = xlLeft
			'end if



		Case "invoice":

			With oExcel.Application.ActiveSheet.PageSetup
				.LeftMargin = 18
				.RightMargin = 18
				.TopMargin = 18
				.BottomMargin = 18
				.HeaderMargin = 18
				.FooterMargin = 18
				.PrintHeadings = False
			end with

			oExcel.cells(11,1).VALUE = "Sold to: " & trim(curClient)
			oExcel.cells(14,6).VALUE = "For Sale and Use in " & trim(curCountry)
			oExcel.cells(datastart-3,2).VALUE = "Prices in US Dollars and for box of 12 units."
		Case "packing":
			oExcel.cells(11,1).VALUE = "Shipped to: " & trim(curClient)
			oExcel.cells(14,6).VALUE = "For Sale and Use in " & trim(curCountry)
			oExcel.Worksheets("Sheet1").Range("F14:F15").Font.Bold = True
	End Select


	oExcel.Worksheets("Sheet1").Range("A11:A12").Font.Bold = True
	oExcel.Worksheets("Sheet1").Range("A15:A16").Font.Bold = True
	oExcel.Worksheets("Sheet1").Range("B11:B11").MergeCells = True
	oExcel.Worksheets("Sheet1").Range("H2:K2").MergeCells = True
	oExcel.ActiveSheet.Range("H2").HorizontalAlignment = xlCenter
	oExcel.Worksheets("Sheet1").Range("H3:K3").MergeCells = True
	oExcel.ActiveSheet.Range("H3").HorizontalAlignment = xlCenter
	oExcel.Columns("J:J").EntireColumn.AutoFit





	oExcel.Range("H6").Select
	'oexcel.ActiveCell.FormulaR1C1 = "=TODAY()"
	oexcel.ActiveCell.FormulaR1C1 = date()
	oexcel.Selection.NumberFormat = "mmmm d, yyyy"
	'oExcel.Range("J6").Select
	'oExcel.Selection.Font.Bold = True
	oExcel.ActiveSheet.Range("H6").Font.Bold = True
	oExcel.Worksheets("Sheet1").Range("H6:K6").MergeCells = True
	oExcel.ActiveSheet.Range("H6").HorizontalAlignment = xlCenter

	currow = datastart


	oExcel.ActiveSheet.Range("A"+trim(cSTR(currow-2))).value="ITEM"
	oExcel.ActiveSheet.Range("A"+trim(cSTR(currow-1))).value="#"

	oExcel.ActiveSheet.Range("B"+trim(cSTR(currow-2))).value="CLIENT"
	oExcel.ActiveSheet.Range("B"+trim(cSTR(currow-1))).value="CODE/DESCRIPTION"
	oExcel.ActiveSheet.Range("C"+trim(cSTR(currow-2))).value="DEMETECH"
	oExcel.ActiveSheet.Range("C"+trim(cSTR(currow-1))).value="CODE"


	oExcel.ActiveSheet.Range("D"+trim(cSTR(currow-1))).value="DESCRIPTION"
	oExcel.ActiveSheet.Range("E"+trim(cSTR(currow-1))).value="CM"
	oExcel.ActiveSheet.Range("F"+trim(cSTR(currow-1))).value="COLOR"
	oExcel.ActiveSheet.Range("G"+trim(cSTR(currow-2))).value="THREAD"
	oExcel.ActiveSheet.Range("G"+trim(cSTR(currow-1))).value="DIAMETER"
	oExcel.ActiveSheet.Range("H"+trim(cSTR(currow-1))).value="MM"
	oExcel.ActiveSheet.Range("I"+trim(cSTR(currow-2))).value="NEEDLE"
	oExcel.ActiveSheet.Range("I"+trim(cSTR(currow-1))).value="CURVATURE"
	oExcel.ActiveSheet.Range("J"+trim(cSTR(currow-1))).value="NEEDLE"
	oExcel.ActiveSheet.Range("K"+trim(cSTR(currow-1))).value="QTY"

	if OutputType = "proforma" then
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-2))).value=""
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-1))).value="PRICE"
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-1))).value="Total"
		oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-1))).value="NeedleCode"




	end if



	if  OutputType = "invoice" then
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-2))).value="USD"
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-1))).value="FOB Miami"
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-1))).value="Total"
	end if

	if OutputType = "packing" then
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-2))).value="Ctn"
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-1))).value="#"
	end if

	if OutputType = "invoice" or OutputType = "packing" or OutputType = "factory" or OutputType = "factorycomp" or OutputType = "factorycomp2" then
		oExcel.ActiveSheet.Range("K"+trim(cSTR(currow-5))).value="Lot#: " & curLot
		oExcel.ActiveSheet.Range("K"+trim(cSTR(currow-4))).value="Mgfr:"
		oExcel.ActiveSheet.Range("K"+trim(cSTR(currow-3))).value="Exp:"
	end if

	if OutputType = "factory"  then
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-2))).value="Cost"
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-1))).value="USD"
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-1))).value="Total"
      'oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-2))).value="Current"
	   oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-2))).Interior.ColorIndex = 15
       oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-1))).value="Cost"
	    oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-1))).Interior.ColorIndex = 15

      oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-2))).value="Total"
	   oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-2))).Interior.ColorIndex = 15
       oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-1))).value="Cost"
	    oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-1))).Interior.ColorIndex = 15



	    oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-2))).value=""
	   oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-2))).Interior.ColorIndex = 15
       oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-1))).value="Barcode"
	    oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-1))).Interior.ColorIndex = 15

'boxit




	   oexcel.Columns("L:N").Select
		oexcel.Selection.NumberFormat = "$#,##0.00"
	end if



	if  OutputType = "factorycomp2" then
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-2))).value="Factory"
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-1))).value="Charged"
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-1))).value="Total"
      oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-2))).value="Fac. Pricelist"
       oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-1))).value="Cost"
	      oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-2))).value="Fac. Cost"
       oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-1))).value="Total"
 	   oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-2))).value="Cost"
       oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-1))).value="Source"
 	   oExcel.ActiveSheet.Range("Q"+trim(cSTR(currow-2))).value="Cost"
       oExcel.ActiveSheet.Range("Q"+trim(cSTR(currow-1))).value="History"


	   oexcel.Columns("L:O").Select
		oexcel.Selection.NumberFormat = "$#,##0.00"
	end if



	if  OutputType = "factorycomp" then
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-2))).value="Cost"
		oExcel.ActiveSheet.Range("L"+trim(cSTR(currow-1))).value="USD"
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-1))).value="Total"
      oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-2))).value="Factory"
       oExcel.ActiveSheet.Range("N"+trim(cSTR(currow-1))).value="Cost"
	      oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-2))).value="Fac. Cost"
       oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-1))).value="Total"
 	   oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-2))).value="Cost"
       oExcel.ActiveSheet.Range("P"+trim(cSTR(currow-1))).value="Source"
 	   oExcel.ActiveSheet.Range("Q"+trim(cSTR(currow-2))).value="Cost"
       oExcel.ActiveSheet.Range("Q"+trim(cSTR(currow-1))).value="History"


	   oexcel.Columns("L:O").Select



			oexcel.Selection.NumberFormat = "$#,##0.00"



	end if





	oexcel.Columns("G:G").Select
	oexcel.Range("G10").Activate
	oexcel.Selection.NumberFormat = "@"

	oexcel.Columns("L:M").Select
	oexcel.Range("L7").Activate
	    if trim(curpricetype) ="E" and outputtype="proforma" then
			'oexcel.Selection.NumberFormat = "#,##0.00 [$�-1]"
			oexcel.Selection.NumberFormat = "[$�-2] #,##0.00"
		else
          oexcel.Selection.NumberFormat = "$#,##0.00"
		end if



	qry="select * from orderdetails where po='" & trim(passedPO)  &"' and (item_no != '') and (code != '') order by CAST (item_no as int) ASC"
	pototqty = 0
	pototcost = 0
	pototvalue = 0
	IndiaspoolFlag = "N"
	poplatetot = 0

	Set oRS = oConn.Execute(qry)

	'do while not oRS.eof
	'presentdemetechcode=oRS.Fields("Code")

		if  not oRS.eof  then
			costlinetot = 0
			faccostlinetot = 0

			lPrintPremMess = "N"

			do while not oRS.eof

				if trim(ucase(curCountry)) = "INDIA"   and instr(1,ucase(oRs.Fields("cm")),"MT") > 1 then

					IndiaspoolFlag = "Y"
				end if

				oExcel.ActiveSheet.Range("A"+trim(cSTR(currow))).value=trim(oRs.Fields("item_no"))
				oExcel.ActiveSheet.Range("B"+trim(cSTR(currow))).value=trim(oRs.Fields("desc"))
				oExcel.ActiveSheet.Range("C"+trim(cSTR(currow))).value=trim(oRs.Fields("code"))
				'response.write(oExcel.ActiveSheet.Range("C"+trim(cSTR(currow-1))).value)
				'response.write(oExcel.ActiveSheet.Range("C"+trim(cSTR(currow))).value)
				presentdemetechcode=trim(oRs.Fields("code"))
				selectneedle(presentdemetechcode)
				'oExcel.ActiveSheet.Range("N"+trim(cSTR(currow))).value=presentdemetechcode
			'	response.Write("hello")
				'marker
				if ucase(right(trim(oRs.Fields("code")),1))="P" then
					lPrintPremMess = "Y"
				end if

				if  trim(oRs.Fields("type")) ="" and OutputType = "proforma" then
					oExcel.ActiveSheet.Range("D"+ trim(cSTR(currow))).Interior.ColorIndex = 6
				else
					oExcel.ActiveSheet.Range("D"+trim(cSTR(currow))).value=trim(oRs.Fields("type"))
				end if

				if  trim(oRs.Fields("cm")) ="" and OutputType = "proforma" then
					oExcel.ActiveSheet.Range("E"+ trim(cSTR(currow))).Interior.ColorIndex = 6
				else
					oExcel.ActiveSheet.Range("E"+trim(cSTR(currow))).value=trim(oRs.Fields("cm"))
				end if

				if  trim(oRs.Fields("color")) ="" and OutputType = "proforma" then
					oExcel.ActiveSheet.Range("F"+ trim(cSTR(currow))).Interior.ColorIndex = 6
				else
					oExcel.ActiveSheet.Range("F"+trim(cSTR(currow))).value=trim(oRs.Fields("color"))
				end if

				if  trim(oRs.Fields("usp")) ="" and OutputType = "proforma" then
					oExcel.ActiveSheet.Range("G"+ trim(cSTR(currow))).Interior.ColorIndex = 6
				else
					oExcel.ActiveSheet.Range("G"+trim(cSTR(currow))).value=trim(oRs.Fields("USP"))
				end if

				if  trim(oRs.Fields("MM")) ="" and OutputType = "proforma"  and not trim(oRs.Fields("CURVATURE")) = "No Needle"  and not ucase(trim(oRs.Fields("CURVATURE"))) = "SPOOL" then
					oExcel.ActiveSheet.Range("H"+ trim(cSTR(currow))).Interior.ColorIndex = 6
				else
					oExcel.ActiveSheet.Range("H"+trim(cSTR(currow))).value=trim(oRs.Fields("MM"))
				end if

				if  trim(oRs.Fields("CURVATURE")) ="" and OutputType = "proforma"  then

						oExcel.ActiveSheet.Range("I"+ trim(cSTR(currow))).Interior.ColorIndex = 6

				else
					oExcel.ActiveSheet.Range("I"+trim(cSTR(currow))).value=trim(oRs.Fields("CURVATURE"))
				end if

				if  trim(oRs.Fields("CROSS_SEC")) ="" and OutputType = "proforma" then
					oExcel.ActiveSheet.Range("J"+ trim(cSTR(currow))).Interior.ColorIndex = 6
				else
					oExcel.ActiveSheet.Range("J"+trim(cSTR(currow))).value=trim(oRs.Fields("CROSS_SEC"))
				end if

				cursamples = (oRs.Fields("samples"))

				if NOT IsNull(cursamples) then

				cursamples = cdbl(cursamples)
				cursamples = 0

				end if
				'if NOT IsNull(cursamples) then

				'end if

				oExcel.ActiveSheet.Range("K"+trim(cSTR(currow))).value= cdbl(oRs.Fields("qty"))'+ trim(oRs.Fields("samples"))





				IF OutputType="proforma" or OutputType="invoice" then

					'add in the distrib multiplie and discount

						x=cdbl(oRs.Fields("qty"))
						y=cdbl(oRs.Fields("intl_price"))
						y = y*cdbl(oRs.Fields("dist"))
						linediscval = y*(cdbl(oRs.Fields("disc"))/100)
						nFinalPrice = round(y - linediscval,2)


					oExcel.ActiveSheet.Range("L"+trim(cSTR(currow))).value=cCur(nFinalPrice)
					'oExcel.ActiveSheet.Range("L"+trim(cSTR(currow))).value=oRs.Fields("intl_PRICE")


						oexcel.Range("L"+trim(cSTR(currow))).Activate
					   	if curpricetype ="E" and outputtype="proforma" then
							'oexcel.Selection.NumberFormat = "#,##0.00 [$�-1]"
							oexcel.Selection.NumberFormat = "[$�-2] #,##0.00"
						else

			                 oexcel.Selection.NumberFormat = "$#,##0.00"
						end if



					'x1=formatnumber(oRs.Fields("QTY"))
					'y1=formatnumber(oRs.Fields("intl_PRICE"))
					'LineTot = x1 * y1

					Linetot = nFinalPrice  * x



					oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value=LineTot
				end if

				if  OutputType="proforma" or OutputType="invoice" then
					if not trim(oRs.Fields("notes")) = "" or cursamples > 0  or cdbl(oRs.Fields("disc")) > 0 then
						if donoteheader="Y" then
							oExcel.ActiveSheet.Range("N22").value="NOTES"
							oExcel.ActiveSheet.Range("N22").Interior.ColorIndex = 15
							oExcel.ActiveSheet.Range("N21").Interior.ColorIndex = 15

							oExcel.ActiveSheet.Range("N22:N23").Select
							oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
							oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
							With oExcel.Selection.Borders(xlEdgeLeft)
								.LineStyle = xlContinuous
								.Weight = xlThin
								.ColorIndex = xlAutomatic
							End With
							With oExcel.Selection.Borders(xlEdgeTop)
								.LineStyle = xlContinuous
								.Weight = xlThin
								.ColorIndex = xlAutomatic
							End With
							With oExcel.Selection.Borders(xlEdgeBottom)
								.LineStyle = xlContinuous
								.Weight = xlThin
								.ColorIndex = xlAutomatic
							End With
							With oExcel.Selection.Borders(xlEdgeRight)
								.LineStyle = xlContinuous
								.Weight = xlThin
								.ColorIndex = xlAutomatic
							End With
							With oexcel.Selection.Borders(xlInsideHorizontal)
						        .LineStyle = 1
						        .Weight = xlThin
						        '.ColorIndex = xlAutomatic
						    End With

						end if
						Notestr = trim(oRs.Fields("notes"))

						if cursamples > 0 then
							notestr =  trim(cstr(cursamples)) + " Free Samples " + notestr
						end if

						if cdbl(oRs.Fields("disc")) > 0  then
							notestr = trim(cstr(oRs.Fields("disc"))) + "% Discount " + Notestr
						end if

						oExcel.ActiveSheet.Range("N"+trim(cSTR(currow))).value=  notestr
						oExcel.Columns("N:N").EntireColumn.AutoFit

						if  OutputType="factory" then
							oExcel.Columns("P:P").EntireColumn.AutoFit
						end if

						donoteheader="N"
					end if
				end if

				pototqty = 	pototqty + trim(oRs.Fields("qty")) '+ trim(oRs.Fields("samples"))
				pototvalue = pototvalue + LineTot

				if ucase(trim(oRs.Fields("CODE")) ="N/A") then
					oExcel.ActiveSheet.Range("A"+ trim(cSTR(currow))+":Z"+trim(cSTR(currow))).font.ColorIndex = 3
					oExcel.ActiveSheet.Range("L"+trim(cSTR(currow))).value="N/A"
					oExcel.ActiveSheet.Range("L"+trim(cSTR(currow))).HorizontalAlignment = xlCenter
				end if

				if  OutputType="factory"   then
					oExcel.ActiveSheet.Range("N"+trim(cSTR(currow))).value=trim(oRs.Fields("cost"))
					oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value =formatcurrency(cdbl(oRs.Fields("cost")) * cdbl(oRs.Fields("qty")))
					pototcost = 	pototcost + cdbl(oRs.Fields("cost")) * cdbl(oRs.Fields("qty"))
					oExcel.ActiveSheet.Range("P"+trim(cSTR(currow))).value	 = "*"+GenBarCodeNum(oExcel.ActiveSheet.Range("C"+trim(cSTR(currow))).value)+"*"
				end if

				if  OutputType="factorycomp" or OutputType="factorycomp2" then
					oExcel.ActiveSheet.Range("L"+trim(cSTR(currow))).value=trim(oRs.Fields("cost"))
					oExcel.ActiveSheet.Range("N"+trim(cSTR(currow))).value=formatcurrency(oRs.Fields("faccost"))
					oExcel.ActiveSheet.Range("P"+trim(cSTR(currow))).value=trim(oRs.Fields("costnotes"))
					oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value =cdbl(oRs.Fields("cost")) * cdbl(oRs.Fields("QTY"))
					oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value =cdbl(oRs.Fields("faccost")) * cdbl(oRs.Fields("QTY"))
					costlinetot = costlinetot + cdbl(oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value)
					faccostlinetot = faccostlinetot+ cdbl(oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value)
					if not  oRs.Fields("faccost") = oRs.Fields("cost") then
						oExcel.ActiveSheet.Range("L"+ trim(cSTR(currow))).Interior.ColorIndex = 6
						oExcel.ActiveSheet.Range("N"+ trim(cSTR(currow))).Interior.ColorIndex = 6
					end if
				end if

			if  OutputType="factorycomp" then
						if cdbl(oRs.Fields("qty")) < 100 and cdbl(oRs.Fields("qty")) > 0  then
							poplatetot = poplatetot + 9.6
 					end if
			end if

				currow = currow+1
				oRS.movenext
		loop
	end if


	if  OutputType="factorycomp2" then
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value = costlinetot
		oExcel.ActiveSheet.Range("K"+trim(cSTR(currow))).HorizontalAlignment = xlCenter
		oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value = faccostlinetot

	end if

	if  OutputType="factory" then
		oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value = formatcurrency(pototcost)
	end if
	if  OutputType="factorycomp" then
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value = costlinetot
		oExcel.ActiveSheet.Range("K"+trim(cSTR(currow))).HorizontalAlignment = xlCenter
		oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value = faccostlinetot

		currow = currow + 1

		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value = poplatetot
		oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value = poplatetot
		currow = currow + 1
		oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value = poplatetot +  costlinetot
		oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).value = poplatetot +  faccostlinetot



	end if

	'Grid the item rows
	if oExcel.ActiveSheet.Range("N22").value="NOTES" or OutputType="factory"    then

		if OutputType="factory" then
			oExcel.Range("A"+trim(cSTR(datastart-2))+":P"+trim(cSTR(currow-1))).Select
		else
			oExcel.Range("A"+trim(cSTR(datastart-2))+":N"+trim(cSTR(currow-1))).Select
				'response.write(oExcel.ActiveSheet.Range("C"+trim(cSTR(currow))).value)
		end if
		'oExcel.Columns("N:N").EntireColumn.Autofit
	else

		if OutputType="factorycomp" or OutputType="factorycomp2" then
			oExcel.Range("A"+trim(cSTR(datastart-2))+":Q"+trim(cSTR(currow-1))).Select
			oExcel.Columns("N:P").EntireColumn.AutoFit

		else
			oExcel.Range("A"+trim(cSTR(datastart-2))+":M"+trim(cSTR(currow-1))).Select
		end if
	end if


	With oexcel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
    With oexcel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
    With oexcel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
    With oexcel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
    With oexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
    With oexcel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = 1
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
	'end of gridding code


	Select Case OutputType
		Case "proforma":
			if mid(passedPO,1,1) = "Q" then
				oExcel.ActiveSheet.Range("H13").value="Price Quote " & passedPO
				oExcel.Worksheets("Sheet1").Range("H13").Font.Bold = True
			else
				oExcel.ActiveSheet.Range("H7").value="PO# " & passedPO
				oExcel.Worksheets("Sheet1").Range("H7").Font.Bold = True
				oExcel.Worksheets("Sheet1").Range("H7").HorizontalAlignment = xlCenter
				oExcel.Worksheets("Sheet1").Range("H7:K7").MergeCells = True

				oExcel.ActiveSheet.Range("E10").value="PROFORMA INVOICE # " & passedPO
				oExcel.Worksheets("Sheet1").Range("E10").Font.Bold = True
				oExcel.Worksheets("Sheet1").Range("E10").Font.Size = 24
				oExcel.Worksheets("Sheet1").Range("E10").HorizontalAlignment = xlCenter
				oExcel.Worksheets("Sheet1").Range("E10:J10").MergeCells = True
				'response.write(oExcel.ActiveSheet.Range("E10").value)

				oExcel.ActiveSheet.Range("E10:J10").Select
					With oExcel.Selection.Borders(xlEdgeLeft)
						.LineStyle = xlContinuous
						.Weight = xlThin
						.ColorIndex = xlAutomatic

					End With
					With oExcel.Selection.Borders(xlEdgeTop)
						.LineStyle = xlContinuous
						.Weight = xlThin
						.ColorIndex = xlAutomatic
					End With
					With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous
						.Weight = xlThin
						.ColorIndex = xlAutomatic
					End With
					With oExcel.Selection.Borders(xlEdgeRight)
						.LineStyle = xlContinuous
						.Weight = xlThin
						.ColorIndex = xlAutomatic
					End With
					With oExcel.Selection.Interior
						.Color=16777062
					End With
					With oExcel.Selection
						.Font.Size = 24
					End With



			end if
		Case "factory" :
			if mid(passedPO,1,1) = "Q" then
				oExcel.ActiveSheet.Range("H13").value="Quote# " & passedPO
			else
				oExcel.ActiveSheet.Range("H13").value="Quote# " & passedPO
			end if

		Case "factorycomp2" :
			if mid(passedPO,1,1) = "Q" then
				oExcel.ActiveSheet.Range("H13").value="Quote# " & passedPO
			else
				oExcel.ActiveSheet.Range("H13").value="Quote# " & passedPO
			end if
		Case "factorycomp" :
			if mid(passedPO,1,1) = "Q" then
				oExcel.ActiveSheet.Range("H13").value="Quote# " & passedPO
			else
				oExcel.ActiveSheet.Range("H13").value="Quote# " & passedPO
			end if
		Case "invoice":
			oExcel.ActiveSheet.Range("H13").value="Invoice# " & passedPO
			oExcel.Worksheets("Sheet1").Range("H13").Font.Bold = True
			oExcel.Worksheets("Sheet1").Range("H13").Font.size = 16
		Case "packing":
			oExcel.ActiveSheet.Range("H13").value="Packing List # " & passedPO
	end select
'marker

	if OutputType="invoice" then
		oExcel.ActiveSheet.Range("F14:K14").Select
	    With oExcel.Selection
			.HorizontalAlignment = xlGeneral
			.VerticalAlignment = xlBottom
			.WrapText = False
			.ShrinkToFit = False
			.MergeCells = True
	    End With
		oExcel.ActiveSheet.Range("F13:K13").Select
		With oExcel.Selection
		  .HorizontalAlignment = xlGeneral
			.VerticalAlignment = xlBottom
			.WrapText = False
			.ShrinkToFit = False
			.MergeCells = True
		End With
	else
		oExcel.ActiveSheet.Range("F13:K13").Select
    	With oExcel.Selection
		  .HorizontalAlignment = xlCenter
    	    .VerticalAlignment = xlBottom
        	.WrapText = False
	        .ShrinkToFit = False
    	    .MergeCells = True
	    End With


	end if


	if not OutputType = "invoice" then
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		oExcel.Selection.Borders(xlInsideVertical).LineStyle = xlNone
	end if


	IF OutputType="proforma" or OutputType="invoice" or OutputType="factory"  or OutputType="packing"  or OutputType="factorycomp" or OutputType="factorycomp2" then

		IF OutputType="proforma" or OutputType="invoice" then
			oExcel.ActiveSheet.Range("I"+trim(cSTR(currow))).value="TOTAL EX-WORKS MIAMI"

		else

			IF OutputType="factorycomp"    then
				oExcel.ActiveSheet.Range("I"+trim(cSTR(currow-2))).value="Order Total"
				oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-1))).font.bold=True
				oExcel.ActiveSheet.Range("M"+trim(cSTR(currow-2))).font.bold=True
				oExcel.ActiveSheet.Range("I"+trim(cSTR(currow-1))).value="Plate Cost"
				oExcel.ActiveSheet.Range("I"+trim(cSTR(currow-1))).font.bold=True
				oExcel.ActiveSheet.Range("I"+trim(cSTR(currow-2))).font.bold=True
				oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-1))).font.bold=True
				oExcel.ActiveSheet.Range("O"+trim(cSTR(currow-2))).font.bold=True

				oExcel.ActiveSheet.Range("K"+trim(cSTR(currow))).font.bold=True
				oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).font.bold=True
				oExcel.ActiveSheet.Range("O"+trim(cSTR(currow))).font.bold=True

				rangestr = "I"+trim(cSTR(currow-2))+":J"+trim(cSTR(currow-2))
				oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
				oexcel.Range(rangestr).Select
				With oexcel.Selection
				  .HorizontalAlignment = xlCenter
				End With


				rangestr = "I"+trim(cSTR(currow-1))+":J"+trim(cSTR(currow-1))
				oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
				oexcel.Range(rangestr).Select
				With oexcel.Selection
				  .HorizontalAlignment = xlCenter
				End With


			end if

			IF not OutputType="factory"  then
				oExcel.ActiveSheet.Range("I"+trim(cSTR(currow))).value="TOTAL"
			else
				'rangestr = "J"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
				'With oexcel.Selection
				 ' .HorizontalAlignment = xlCenter
				'End With
			end if

		end if

		rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
		oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
		oexcel.Range(rangestr).Select
		With oexcel.Selection
		  .HorizontalAlignment = xlCenter
		End With
		oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True
	end if

	'IF cOutputType="FACTORY" then
	'	oExcel.ActiveSheet.Range("L19").value="COST"
	'	oExcel.ActiveSheet.Range("M19").value="TOTAL"
	'end if

	'oExcel.Range("A18:P19").Select
	'oExcel.Selection.Interior.ColorIndex = 15
	'oExcel.Selection.Font.Bold = True












		oExcel.Range(("A"+trim(cSTR(currow)))).Select
		oExcel.Selection.Font.Bold = True
'marker
		oExcel.Range(("A"+trim(cSTR(currow+1)))).Select
		oExcel.Selection.Font.Bold = True

		oExcel.Range(("K"+trim(cSTR(currow)))).Select
		oExcel.Selection.Font.Bold = True
		oExcel.Selection.HorizontalAlignment = xlCenter


		oExcel.ActiveSheet.Range("K"+trim(cSTR(currow))).value =	pototqty
		if OutputType="invoice" or OutputType="proforma" then
			oExcel.ActiveSheet.Range("M"+trim(cSTR(currow))).value = pototvalue
		end if


		if not outputtype ="factorycomp" or OutputType="factorycomp2" then
			rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
		else
			rangestr = "I"+trim(cSTR(currow))+":O"+trim(cSTR(currow))
		end if

		if outputtype ="factory" then
			rangestr = "I"+trim(cSTR(currow))+":P"+trim(cSTR(currow))
		end if

		'rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		'oExcel.Selection.Borders(xlInsideVertical).LineStyle = xlNone
		oExcel.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous




	oexcel.ActiveSheet.PageSetup.PrintArea = ""
   With oexcel.ActiveSheet.PageSetup
      .Orientation = xlLandscape
  End With

	oexcel.Columns("A:A").ColumnWidth = 4.57


	rangestr="G22:"+"G"+trim(cSTR(currow-1))
	oexcel.Range(rangestr).Select
    With oexcel.Selection
        .HorizontalAlignment = xlCenter
    End With
	rangestr="H22:"+"H"+trim(cSTR(currow-1))
    oexcel.Range(rangestr).Select
    With oexcel.Selection
        .HorizontalAlignment = xlCenter
    End With
	rangestr="K22:"+"K"+trim(cSTR(currow-1))
    oexcel.Range(rangestr).Select
    With oexcel.Selection
        .HorizontalAlignment = xlCenter
    End With
	'oExcel.Columns("B:B").EntireColumn.AutoFit
	rangestr="A22:"+"A"+trim(cSTR(currow-1))

	oexcel.Range(rangestr).Select
    With oexcel.Selection
        .HorizontalAlignment = xlCenter
    End With



	 if OutputType = "proforma" then
			oExcel.ActiveSheet.Range("I"+ trim(cSTR(currow))+":M"+trim(cSTR(currow))).Interior.ColorIndex = 6
			'oExcel.Selection.Font.Bold = .t.
	end if


	nExtraLines = 0

	if OutputType = "proforma"  or  OutputType = "invoice" then

		if not cdbl(curdiscount)=0 then
			nExtraLines =  nExtraLines + 1
			dodiscount()
		end if
		if not cdbl(curadjust)=0 then
			nExtraLines =  nExtraLines + 1
			doadjust()
		end if
		if not cdbl(curfreight)=0 then
			nExtraLines =  nExtraLines + 1
			dofreight()
		end if
		if not cdbl(curadjust)=0 or not cdbl(curdiscount)=0 or not cdbl(curfreight)=0 then
			donettotal()
		end if
		if nExtraLines > 0 then
			nExtraLines = nExtraLines + 1
		end if
	end if

	if OutputType = "factory" then

			nExtraLines =  nExtraLines + 1
			dofacdiscount()
			donetfactotal()

	end if

	IF OutputType = "proforma" or outputype="invoice" then



		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).value="*Minimum order 50 boxes per code."
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).font.bold=True

		'if lPrintPremMess = "Y" then
			'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).value="*DemeTech PLUS line (Designated with a P at the end of the code) Delivery time is 4-12 weeks"
			'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).font.bold=True
		'end if
		if len(pay_mess) = 0 then
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).value="*Production Time: 4-12 Weeks"
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).font.bold=True
			 oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).value="*Minimum order 50 boxes per code"
			 oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).font.bold=True
			 oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).value="*50% deposit with order, balance prior to shipment"
			 oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).font.bold=True
			'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).value="*Payment with order"
		else
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).value="*Production Time: 4-12 Weeks"
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).font.bold=True
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).value=" *Minimum order 50 boxes per code. "
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).font.bold=True
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).value= pay_mess
			oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).font.bold=True
		end if


	' "First American Bank
       ' 1650 Louis Avenue
       ' Elk Grove Village, IL 60007, USA,
       ' ABA no. 071-922-777,
       ' for credit to the Trade Finance Department
       ' account no. 07811228301
       ' Reference:  Demetech"



		'if ucase(trim(curCountry))= "RUSSIA" and ucase(trim(curClient)) = "HOSPITAL INTENSIVE" then



		'oExcel.cells(13,1).VALUE = "Dorognaya 60,				                                  "
		'oExcel.cells(14,1).VALUE = "Corp. 1               		                             "
		'oExcel.cells(15,1).VALUE = "Moscow                   	                                  "
		'oExcel.cells(16,1).VALUE = "RUSSIA"

		'variable="*Payment 90 days from Invoice."
		'Dibujar()

		if ucase(trim(curCountry))= "JORDAN" and ucase(trim(curClient)) = "AL HILAL DRUG STORE" then


		oExcel.cells(13,1).VALUE = "Sarh Al - Shaheed Street,                            "
		oExcel.cells(14,1).VALUE = "P.O. Box 203                        "
		oExcel.cells(15,1).VALUE = "Ammam  11947                            "
		oExcel.cells(16,1).VALUE = "JORDAN"

		variable="*Payment 90 days from Invoice."
		Dibujar()

		elseif ucase(trim(curCountry))= "SOUTH AFRICA" and ucase(trim(curClient)) = "BAROQUE MEDICAL (PTY) LTD." then


		oExcel.cells(13,1).VALUE = "12 Rivonia Road,                            "
		oExcel.cells(14,1).VALUE = "Illovo, P.O. Box 785"
		oExcel.cells(15,1).VALUE = "Northlands, 2116                            "
		oExcel.cells(16,1).VALUE = "SOUTH AFRICA"


		variable="*Payment 90 days from Invoice."
		Dibujar()

		elseif ucase(trim(curCountry))= "HONDURAS" and ucase(trim(curClient)) = "HEALTHCARE PRODUCTS CENTROAMERICA, S. DE R.L." then


		'oExcel.cells(13,1).VALUE = "12 Rivonia Road,                            "
		'oExcel.cells(14,1).VALUE = "Illovo, P.O. Box 785"
		'oExcel.cells(15,1).VALUE = "Northlands, 2116                            "
		'oExcel.cells(16,1).VALUE = "SOUTH AFRICA"


		variable="*Payment 60 days from Invoice."
		Dibujar()


		elseif ucase(trim(curCountry))= "ECUADOR" and ucase(trim(curClient)) = "REPRESENTACIONES MEDICAS JF" then


		'oExcel.cells(13,1).VALUE = "12 Rivonia Road,                            "
		'oExcel.cells(14,1).VALUE = "Illovo, P.O. Box 785"
		'oExcel.cells(15,1).VALUE = "Northlands, 2116                            "
		'oExcel.cells(16,1).VALUE = "SOUTH AFRICA"


		variable="*Payment 60 days from Invoice."
		Dibujar()





		elseif ucase(trim(curCountry))= "LEBANON" and ucase(trim(curClient)) = "ZOUHEIR DANDAN EST." then


		'Zouheir Dandan Est.
'Mohamad Dandan

		'Dandan Building, 5th Floor,
			'Mazar's Street,
'P.O. Box 11-7545, Riad el Solh
'Berirut  11072240
'LEBANON

		oExcel.cells(13,1).VALUE = "Dandan Building, 5th Floor,                              "
		oExcel.cells(14,1).VALUE = "Mazar's Street                                           "
		oExcel.cells(15,1).VALUE = "P.O. Box 11-7545, Riad el Solh                           "
		oExcel.cells(16,1).VALUE = "Berirut  11072240                                        "
		oExcel.cells(17,1).VALUE = "LEBANON                                                  "

		variable="*Payment 90 days from Invoice."
		Dibujar()

        elseif ucase(trim(curCountry))= "VIETNAM" and ucase(trim(curClient)) = "GOLDEN GATE JSC" then  'Golden Gate JSC


       ' #236//29/18 Dien Bien Phu St.,
'Ward 17, Binh Thanh Dist.,
'Hochiminh City,
'VIETNAM


		oExcel.cells(13,1).VALUE = " #236//29/18 Dien Bien Phu St.,                                         "
		oExcel.cells(14,1).VALUE = "Ward 17, Binh Thanh Dist.,                                              "
		oExcel.cells(15,1).VALUE = "Hochiminh City,                                                         "
		oExcel.cells(16,1).VALUE = "VIETNAM"

		variable="*Payment 30 days from Invoice."
		Dibujar()

		elseif (ucase(trim(curCountry))= "GREECE" or ucase(trim(curCountry))= "GREECE") and ucase(trim(curClient)) = "PASCAL STROUZA S.A." then  'PASCAL

		oExcel.cells(13,1).VALUE = " 36. Kosti Palama Street,                                         "
		oExcel.cells(14,1).VALUE = "N. Chalkodona GR-143-43                                              "
		oExcel.cells(15,1).VALUE = "GREECE                                                        "
		'oExcel.cells(16,1).VALUE = "SPAIN"

		variable="*Payment 90 days from Invoice."
		Dibujar()

		elseif (ucase(trim(curCountry))= "SPAIN" or ucase(trim(curCountry))= "SPAIN" or ucase(trim(curCountry))= "ESPA�A") and ucase(trim(curClient)) = "SANGUESA SA" then  'SANGUESA S A

		oExcel.cells(13,1).VALUE = " Pasaje Mercuri 4,                                         "
		oExcel.cells(14,1).VALUE = "08940 Cornella De Llobregat                                              "
		oExcel.cells(15,1).VALUE = "Barcelona,                                                         "
		oExcel.cells(16,1).VALUE = "SPAIN"

		variable="*Payment 60 days from Invoice."
		Dibujar()
		elseif (ucase(trim(curCountry))= "EGYPT" or ucase(trim(curCountry))= "EGYPT") and ucase(trim(curClient)) = "BM EGYPT MEDICAL AND SCIENTFIC EQUIPMENT SAE" then  'BM-Egypt


       ' #236//29/18 Dien Bien Phu St.,
'Ward 17, Binh Thanh Dist.,
'Hochiminh City,
'VIETNAM


		oExcel.cells(13,1).VALUE = " 23 Iran Street, Dokki,                                          "
		oExcel.cells(14,1).VALUE = "Giza                                              "
		oExcel.cells(15,1).VALUE = "EGYPT                                                         "
		'oExcel.cells(16,1).VALUE = "EGYPT"

		variable="*Payment 30 days from Invoice."
		Dibujar()


		elseif (ucase(trim(curCountry))= "MEXICO" or ucase(trim(curCountry))= "MEXICO") and ucase(trim(curClient)) = "SAVI DISTRIBUCIONES S.A. DE C.V." then



		oExcel.cells(13,1).VALUE = "(Alternavida, S.A. de C. V.)                                         "
		oExcel.cells(14,1).VALUE = "Av. Magnocentro No.11 Pisos                                               "
		oExcel.cells(15,1).VALUE = "2,3 y 5 Col.Centro Urb Interlo                                                        "
		oExcel.cells(16,1).VALUE = "Huixquilucan, MEXICO"

		variable="*Payment 60 days from Invoice."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).value="*By accepting this Proforma Invoice, the purchaser agrees to pay for these medical devices on the above mentioned payment terms."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).font.bold=True

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).value="*Specialty Items may have longer production times."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).font.bold=True


		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+8-nExtraLines))).value=""
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+8-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).value="Regions Bank"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).value="ACCT: # 0198073737"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).value="ABA: 063104668"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).value="8900 SW 107 Ave"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).value="Miami, FL 33176"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).value="SWIFT: UPNBUS44MIA"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).value="Phone: (305) 596-3697"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).value=trackstr
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).value="Deposits are Non-refundable, all sales are finals, no returns, confirmed Purchase Orders can not be modified or canceled"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).value="All sales are finals, no returns"
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.size=60
        ' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).value="Confirmed Purchase Orders can not be modified or canceled"
        ' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).font.size=26
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).value=""
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.bold=True
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).value=""
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).font.bold=True
		Dibujar2()

		elseif (ucase(trim(curCountry))= "HUNGARY" or ucase(trim(curCountry))= "HUNGARY") and ucase(trim(curClient)) = "GOLD MEDICAL" then  'Golden Medical

		'oExcel.cells(13,1).VALUE = "2370 Dabas Dinny�s,                                         "
		'oExcel.cells(14,1).VALUE = "Lajos u. 12,                                              "
		'oExcel.cells(15,1).VALUE = "HUNGARY"

		'oExcel.cells(13,1).VALUE = "Dinnyes Lajos Street 12, Dabas                                         "
		 oExcel.cells(13,1).VALUE = "1221 Budapest Hasad�k utca 22/b                                        "

		'oExcel.cells(14,1).VALUE = "Lajos u. 12,                                              "
		oExcel.cells(15,1).VALUE = "HUNGARY"

		variable="*Payment 30 days from Invoice."
		Dibujar()

		elseif ucase(trim(curCountry))= "INDONESIA" and ucase(trim(curClient)) = "PT. ENSEVAL MEDIKA PRIMA" then
		'Pt. Enseval Medika Prima
		'Jalan Pulo Lentut No#10,
		'Kawasan Industri Pulo Gadung
		'Jakarta 13920
		'INDONESIA

		oExcel.cells(13,1).VALUE = "Jalan Pulo Lentut No#10,                                          "
		oExcel.cells(14,1).VALUE = "Kawasan Industri Pulo Gadung                                      "
		oExcel.cells(15,1).VALUE = "Jakarta 13920                                                     "
		oExcel.cells(16,1).VALUE = "INDONESIA"

		variable="*Payment terms: Net 60 days."
		Dibujar()

		elseif ucase(trim(curCountry))= "SAUDI ARABIA" and ucase(trim(curClient)) = "JAMJOOM MEDICAL INDUSTRIES COMPANY" then


		oExcel.cells(13,1).VALUE = "Industrial Area, Phase #4,                           "
		oExcel.cells(14,1).VALUE = "Street #403,                                         "
		oExcel.cells(15,1).VALUE = "Makkah Highway Road                                  "
		oExcel.cells(16,1).VALUE = "Jeddah                                               "
		oExcel.cells(17,1).VALUE = "SAUDI ARABIA"

		variable="*Payment 90 days from Invoice."
		Dibujar()


		elseif ucase(trim(curCountry))= "SAUDI ARABIA" and ucase(trim(curClient)) = "HUMEIDAN & KHOWAITER CO." then


		oExcel.cells(13,1).VALUE = "P.O. Box 288                          "
		oExcel.cells(14,1).VALUE = "Building 70,                          "
		oExcel.cells(15,1).VALUE = "Al Washeem Street                     "
		oExcel.cells(16,1).VALUE = "11411 Riyadh"
		oExcel.cells(17,1).VALUE = "SAUDI ARABIA"

		variable="*Payment 60 days from Invoice."
		Dibujar()

		elseif ucase(trim(curCountry))= "POLAND" and ucase(trim(curClient)) = "TOP MED" then



		oExcel.cells(13,1).VALUE = "Ul. Sczanieckiej 7/7,                     "
		oExcel.cells(14,1).VALUE = "93-342 Lodz                                       "
		oExcel.cells(15,1).VALUE = "POLAND                                            "


		variable="*Payment 30 days from Invoice."
		Dibujar()

     elseif ucase(trim(curCountry))= "KENYA" and ucase(trim(curClient)) = "STATIM PHARMACEUTICALS LTD" then
        variable="*Payment 30 days from Invoice."
		Dibujar()
	 elseif ucase(trim(curCountry))= "NEW ZEALAND" and ucase(trim(curClient)) = "MEDENT MEDICAL" then



	'	oExcel.cells(13,1).VALUE = "Ul. Sczanieckiej 7/7,                     "
	'	oExcel.cells(14,1).VALUE = "93-342 Lodz                                       "
	'	oExcel.cells(15,1).VALUE = "POLAND                                            "


		variable="*Payment 30 days from Invoice."
		Dibujar()



		'elseif ucase(trim(curCountry))= "EGYPT" and ucase(trim(curClient)) = "I.M. TRADING & MARKETING CO." then




		'oExcel.cells(13,1).VALUE = "12 Al Aziz Bellah Street                                            "
		'oExcel.cells(14,1).VALUE = "In Front of Naser Social Bank,                                        "
		'oExcel.cells(15,1).VALUE = "Al Zaitoun - Heliopolis"
		'oExcel.cells(16,1).VALUE = "Cairo"
		'oExcel.cells(17,1).VALUE = "EGYPT"

		'variable="*Payment 90 days from Invoice."
		'Dibujar()

		'elseif ucase(trim(curCountry))= "SRI LANKA" and ucase(trim(curClient)) = "MARKSS HLC (PRIVATE) LIMITED" then

		'Markss HLC (Private) LImited




		'oExcel.cells(13,1).VALUE = "153/3 Nawala Road,                 "
		'oExcel.cells(14,1).VALUE = "Narahenpita,              		                                       "
		'oExcel.cells(15,1).VALUE = "Colombo 05,                        "
		'oExcel.cells(16,1).VALUE = "SRI LANKA"

		'variable="*Payment 60 days from Invoice."
		'Dibujar()

		elseif ucase(trim(curCountry))= "VIETNAM" and ucase(trim(curClient)) = "MINH PHUONG TECHNICAL MEDICAL EQUIPMENT TRADING COMPANY, LTD." then

		oExcel.cells(11,1).VALUE = "Quote to: Minh Phuong Technical Medical Equipment"
		oExcel.cells(12,1).VALUE = "Trading Company, Ltd."
		oExcel.cells(13,1).VALUE = "Doan Phuong                       "
		oExcel.cells(14,1).VALUE = "105E2 Phuong Mai,                 "
		oExcel.cells(15,1).VALUE = "Dong Da              		   	  "
		oExcel.cells(16,1).VALUE = "Hanoi 65201                       "
		oExcel.cells(17,1).VALUE = "VIETNAM"

		variable="*Payment 60 days from Invoice."
		Dibujar()

		elseif ucase(trim(curCountry))= "COLOMBIA" and ucase(trim(curClient)) = "VESALIUS PHARMA S.A.S." then


		oExcel.cells(13,1).VALUE = "Cra 21 No169-76 Of. 208"
		oExcel.cells(14,1).VALUE = "Bogota, COLOMBIA"
		'oExcel.cells(15,1).VALUE = "Umraniye Istanbul 81230"
		'oExcel.cells(16,1).VALUE = "COLOMBIA"

		variable="*Payment 30 days from Invoice."
		Dibujar()


		elseif ucase(trim(curCountry))= "NETHERLANDS" and ucase(trim(curClient)) = "QRS HEALTHECARE B.V." then


		oExcel.cells(13,1).VALUE = "Post Bus 390					"
		oExcel.cells(14,1).VALUE = "5340 AJ Oss Eindhoven			"
		oExcel.cells(15,1).VALUE = "NETHERLANDS						"


		variable="*Payment 60 days from Invoice."
		Dibujar()


		elseif ucase(trim(curCountry))= "AUSTRALIA" and ucase(trim(curClient)) = "EBOS GROUP LIMITED" then
        response.write("TEST")

		'oExcel.cells(13,1).VALUE = "Post Bus 390					"
		'oExcel.cells(14,1).VALUE = "5340 AJ Oss Eindhoven			"
		'oExcel.cells(15,1).VALUE = "NETHERLANDS						"


		variable="*Payment 30 days from Invoice."
		'	msgbox(variable)
		Dibujar()

		elseif ucase(trim(curCountry))= "TURKEY" and ucase(trim(curClient)) = "MARMED MEDIKAL" then
        response.write("TEST")

		'oExcel.cells(13,1).VALUE = "Post Bus 390					"
		'oExcel.cells(14,1).VALUE = "5340 AJ Oss Eindhoven			"
		'oExcel.cells(15,1).VALUE = "NETHERLANDS						"


		variable="*Payment 120 days from Invoice."
		'	msgbox(variable)
		Dibujar()
		elseif ucase(trim(curCountry))= "NEW ZEALAND" and ucase(trim(curClient)) = "EBOS HEALTHCARE" then
        response.write("TEST")

		'oExcel.cells(13,1).VALUE = "Post Bus 390					"
		'oExcel.cells(14,1).VALUE = "5340 AJ Oss Eindhoven			"
		'oExcel.cells(15,1).VALUE = "NETHERLANDS						"


		variable="*Payment 30 days from Invoice."
		'	msgbox(variable)
		Dibujar()

		elseif ucase(trim(curCountry))= "NEW ZEALAND" and ucase(trim(curClient)) = "AMTECH MEDICAL" then
				response.write("TEST")

		variable="*Payment 30 days from Invoice."
		'	msgbox(variable)
		Dibujar()

		elseif ucase(trim(curCountry))= "DOMINICAN REPUBLIC" and ucase(trim(curClient)) = "GENERICOS DEL CARIBE, S.R. L" then

		oExcel.cells(13,1).VALUE = "Calle Club Activo 20-30 #118                             "
		oExcel.cells(14,1).VALUE = "Santo Domingo"
		oExcel.cells(15,1).VALUE = "DOMINICAN REPUBLIC                            "
		'oExcel.cells(16,1).VALUE = "SOUTH AFRICA"

		variable="*Payment 30 days from Invoice."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).value="*By accepting this Proforma Invoice, the purchaser agrees to pay for these medical devices on the above mentioned payment terms."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).font.bold=True
    	oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).value="*Specialty Items may have longer production times."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).font.bold=True

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+8-nExtraLines))).value=""
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+8-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).value="Regions Bank"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).value="ACCT: # 0198073737"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).value="ABA: 063104668"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).value="8900 SW 107 Ave"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).value="Miami, FL 33176"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).value="SWIFT: UPNBUS44MIA"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).value="Phone: (305) 596-3697"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).value=trackstr
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).value="Deposits are Non-refundable, all sales are finals, no returns, confirmed Purchase Orders can not be modified or canceled"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).value="All sales are finals, no returns"
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.size=60
        ' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).value="Confirmed Purchase Orders can not be modified or canceled"
        ' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).font.size=26
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).value=""
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.bold=True
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).value=""
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).font.bold=True
		Dibujar2()



		else

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).value="*By accepting this Proforma Invoice, the purchaser agrees to pay for these medical devices on the above mentioned payment terms."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).font.bold=True

	    oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).value="*Specialty Items may have longer production times."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).font.bold=True



	''	oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).value=""
	''	oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).value="Regions Bank"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).value="ACCT: # 0198073737"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).value="ABA: 063104668"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).value="8900 SW 107 Ave"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).value="Miami, FL 33176"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).value="SWIFT: UPNBUS44MIA"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).value="Phone: (305) 596-3697"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+15-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).value=trackstr
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).value="Deposits are Non-refundable, all sales are finals, no returns, confirmed Purchase Orders can not be modified or canceled"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).value="All sales are finals, no returns"
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.size=60
        ' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).value="Confirmed Purchase Orders can not be modified or canceled"
        ' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+19-nExtraLines))).font.size=26
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).value=""
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).font.bold=True
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).value=""
        oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.bold=True
		Dibujar2()
		end if

	end if



	'if ucase(trim(curCountry))="USA" or (ucase(trim(curCountry))="SPAIN" and ucase(trim(curClient))="SANGUESA SA") or (ucase(trim(curClient))="HELAL TEB ")then
	'if NOT mid(passedPO,1,1) = "Q" then
				If ucase(trim(curCountry))="USA" Then
					oexcel.Range("F21").Select
					oexcel.Selection.EntireColumn.Insert
					OEXCEL.Worksheets("Sheet1").Range("F20:F21").Font.Bold = True
					oExcel.cells(21,6).VALUE =""
					oExcel.cells(22,6).VALUE="INCHES"
				END IF

	'	IF ucase(trim(curCountry))="USA" or (ucase(trim(curCountry))="SPAIN" and ucase(trim(curClient))="SANGUESA SA") or (ucase(trim(curClient))="HELAL TEB ") then
					oexcel.Range("I20").Select
					oexcel.Selection.EntireColumn.Insert
					OEXCEL.Worksheets("Sheet1").Range("I20:I21").Font.Bold = True
					oExcel.cells(21,9).VALUE ="NEEDLE"
					oExcel.cells(22,9).VALUE="CODE"
					oexcel.Range("R19").Select
					oexcel.Selection.EntireColumn.Insert
					OEXCEL.Worksheets("Sheet1").Range("R18:R19").Font.Bold = True

					IF OutputType = "factory" then
						oExcel.cells(18,18).VALUE ="SULZE"
						oExcel.cells(19,18).VALUE="CODE"
					end if

					oExcel.ActiveSheet.Range("N"+ trim(cSTR(datastart-2))+":O"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15
					oEXCEL.Worksheets("Sheet1").Range("P18:S19").Font.Bold = True


					IF OutputType = "proforma" or outputype="invoice" then
						datastart=23
					else
						datastart=23
						'datastart=25
					end if


					TotalNoRows=oExcel.ActiveSheet.UsedRange.Rows.Count
					FacCoderowcnt = datastart
					'response.write(datastart)
		         response.write("-"&TotalNoRows&"Comienza"&FacCoderowcnt)
					DO WHILE FacCoderowcnt <= TotalNoRows-1

											'response.write("@"&faccoderowcnt&"-"&oExcel.Cells(FacCoderowcnt, 2).Value)
											'response.write(TotalNoRows)
								dim Vlo
								'response.write(faccoderowcnt)
                                Vlo="@"&faccoderowcnt&"-"&oExcel.Cells(FacCoderowcnt, 2).Value
								'response.write(Vlo)
								qry="select * from newethicon where upper(ecode)=upper('" + Vlo + "')"
								'qry="select * from newethicon where upper(ecode)=upper('" + oExcel.Cells(FacCoderowcnt, 2).Value + "')"
								''response.write("-"&qry&"_")
								Set oRSFacCodes = oConn.Execute(qry)
				'              >>>----------------------------------------------------------------------------------------------------> Asignar valores
								Dim ProdCode
								Dim ConCurvature
								Dim ConCross
								Dim ConMM
								Dim ConCode
								Dim mm
								Dim Curvature
								Dim Cross
								'Response.write(oExcel.cells(23,2).VALUE)
								'Response.write(oExcel.cells(23,8).VALUE)
								'Response.write(oExcel.cells(23,3).VALUE & "-")
								'Response.write(oExcel.cells(23,4).VALUE & "-")
								'Response.write(oExcel.cells(22,5).VALUE & "-")
								if oExcel.cells(22,2).VALUE ="CODE" Then
								ConCode= 2
								elseif oExcel.cells(22,3).VALUE ="CODE" Then
								ConCode= 3
								elseif oExcel.cells(22,4).VALUE ="CODE" Then
								ConCode= 4
								End If
								ProdCode = oExcel.Cells(FacCoderowcnt, ConCode).Value
								'Response.write("-" & ProdCode)
						'Response.write("--"&ncode&"-"&ProdCode&"-"&demecodes)
								 'Response.write(oExcel.cells(22,2).VALUE & "-")
								 'Response.write(oExcel.cells(22,3).VALUE & "-")
								'Response.write(oExcel.cells(22,4).VALUE & "-")
								' Response.write(oExcel.cells(22,8).VALUE & "-")
								' Response.write(oExcel.cells(22,9).VALUE & "-")
								' Response.write(oExcel.cells(22,10).VALUE & "-")
									' Response.write(oExcel.cells(22,11).VALUE & "-")
										' Response.write(oExcel.cells(22,12).VALUE & "-")
											' Response.write(oExcel.cells(22,13).VALUE & "-")
												' Response.write(oExcel.cells(22,14).VALUE & "-")
													' Response.write(oExcel.cells(22,15).VALUE & "-")
														' Response.write(oExcel.cells(22,16).VALUE & "-")
								If ucase(trim(curCountry))="USA" Then
									if oExcel.cells(22,7).VALUE ="MM" Then
										ConMM= 7
										elseif oExcel.cells(22,8).VALUE ="MM" Then
										ConMM= 8
										elseif oExcel.cells(22,9).VALUE ="MM" Then
										ConMM= 9
										elseif oExcel.cells(22,10).VALUE ="MM" Then
										ConMM= 10
										elseif oExcel.cells(22,11).VALUE ="MM" Then
										ConMM= 11
									End if
									if oExcel.cells(22,11).VALUE ="CURVATURE" Then
										ConCurvature= 11
										elseif oExcel.cells(22,12).VALUE ="CURVATURE" Then
										ConCurvature= 12
										elseif oExcel.cells(22,13).VALUE ="CURVATURE" Then
										ConCurvature= 13
									End if
									if oExcel.cells(22,14).VALUE ="NEEDLE" Then
										ConCross= 14
										elseif oExcel.cells(22,12).VALUE ="NEEDLE" Then
										ConCross= 12
										elseif oExcel.cells(22,15).VALUE ="NEEDLE" Then
										ConCross= 15
										elseif oExcel.cells(22,16).VALUE ="NEEDLE" Then
										ConCross= 16
								    End if
								else
									if oExcel.cells(22,5).VALUE ="MM" Then
											ConMM= 5
											elseif oExcel.cells(22,6).VALUE ="MM" Then
											ConMM= 6
											elseif oExcel.cells(22,7).VALUE ="MM" Then
											ConMM= 7
											elseif oExcel.cells(22,8).VALUE ="MM" Then
											ConMM= 8
									End if
									if oExcel.cells(22,8).VALUE ="CURVATURE" Then
										ConCurvature= 8
										elseif oExcel.cells(22,9).VALUE ="CURVATURE" Then
										ConCurvature= 9
										elseif oExcel.cells(22,10).VALUE ="CURVATURE" Then
										ConCurvature= 10
									End if
									if oExcel.cells(22,9).VALUE ="NEEDLE" Then
										ConCross= 9
										elseif oExcel.cells(22,10).VALUE ="NEEDLE" Then
										ConCross= 10
										elseif oExcel.cells(22,11).VALUE ="NEEDLE" Then
										ConCross= 11
								    End if
								End If
								  mm = oExcel.Cells(FacCoderowcnt, ConMM).Value
								  Curvature = oExcel.Cells(FacCoderowcnt, ConCurvature).Value
						  		  Cross = oExcel.Cells(FacCoderowcnt, ConCross).Value
										' Response.write("-" & mm)
										' Response.write("-" & Curvature)
										'Response.write("-" & Cross)


							' Response.write("-" & Cross)

								Dim a
				                Dim b
								Dim var
								Dim CmmX
								CmmX= mm
								Dim CmmX2
								Dim cNdlLookup
								Dim demecodes
								Dim cNdlImage
								Dim needleCode
								Dim Code
								Dim qry2
								Dim qry3
								Dim qry4
								Dim ncode
								'Response.write("--"&ProdCode)
								If InStr(1,CmmX,",") > 0 Then
									CmmX2= Replace(Cmmx,",","")
									CmmX2="0" + CmmX2
								ElseIf InStr(1,CmmX,".") > 0 Then
									CmmX2= Replace(Cmmx,".","")
									CmmX2="0" + CmmX2
								Else
									CmmX2=CmmX
								End If
							  ' Response.write(CmmX2)

								 qry2="SELECT * FROM needle_image where type='NC' and descripcion='" + Curvature + "'"
								Response.write("_"&qry2)
								 Set NeedleCommand = oConn.Execute(qry2)
								 	If Not NeedleCommand.eof Then
									 demecodes = demecodes & CmmX2
							     	cNdlImage= CmmX2 & "_"
									 needleCode = NeedleCommand.Fields("needlecode")
								'   Response.write(NeedleCommand.Fields("needlecode"))
								     code = NeedleCommand.Fields("code")
								     demecodes = demecodes + needleCode
								     cNdlImage = cNdlImage + code  & "_"
									 'Response.write(cNdlImage)
							     	var=""
									qry3="SELECT * FROM needle_image where type='CS' and descripcion='" + Cross + "'"


									' Response.write(qry3)
									Set NeedleCommand2 = oConn.Execute(qry3)
									If Not NeedleCommand2.eof Then
											needleCode = NeedleCommand2.Fields("needlecode")
											code = NeedleCommand2.Fields("code")

											demecodes  = demecodes + needleCode
											cNdlImage = cNdlImage + needleCode
											' Response.write(demecodes & "-")

												If InStr(1,ProdCode,",") > 0 AND NOT InStr(1,ProdCode,"-") > 0 Then
												If InStr(1,demecodes,"0") > 0 AND NOT InStr(2,demecodes,"0") > 0 Then
													a=Split(demecodes,"0")
													b=UBound(a)
													if UBound(a) > 2 Then
														for i = 1 to 1
														demecodes = "0" & "," & a(1) & "0"
														next
													elseif UBound(a) = 2 Then
														for i = 1 to 1
														demecodes = "0" & "," & a(1) & "0"
														next
													else
														for i = 0 to 0
														demecodes = "0" & "," & a(0) & "0"
														next
													end if
												else
														demecodes="0" & "," & demecodes
												End If
											End If

										'	Response.write(demecodes)	  'Ta bien con codigos sin coma -.-
											Dim MyStr
											MyStr=Right(prodcode,1)
											If MyStr = "P" Then
												demecodes=demecodes & "P"
											elseif MyStr = "M" Then
												demecodes=demecodes & "M"
											else
												MyStr=Right(prodcode,2)
												If MyStr = "VB" Then
													demecodes=demecodes & "VB"
												else
													demecodes=""
												End if
											End If

										'Response.write("--"&ProdCode)

									'	Response.write(demecodes)
										If Not demecodes= "" Then
										'	Response.write(mm)
										'	Response.write(Curvature)
										'	Response.write(Cross)
										'	Response.write(demecodes)
										   qry4="SELECT * FROM needles1 where mm='" & mm & "' and curvature='" & Curvature & "' and Cross_sec='" & Cross & "' and demecode = '" & demecodes & "'"
										'	Response.write(qry4)
											 Set NeedleCommand3 = oConn.Execute(qry4)
											 If Not NeedleCommand3.eof Then
											 ncode=NeedleCommand3.Fields("dneedle")
											 End If
										Else
											ncode=""
										End if
										oExcel.Cells(FacCoderowcnt, 9).Value = trim(ncode)
									'Response.write("--"&ncode&"-"&ProdCode&"-"&demecodes)
									End IF
								End If
										demecodes=""
										ncode=""




							if not oRSFacCodes.eof then
								IF OutputType = "factory" then
									oExcel.Cells(FacCoderowcnt, 18).Value = oRSFacCodes.Fields("ethiconnee")
								end if

							'	oExcel.Cells(FacCoderowcnt, 9).Value = trim(oRSFacCodes.Fields("intl_ecode"))

							end if
						if ucase(trim(curCountry))="USA"THEN
							select case trim(oExcel.Cells(FacCoderowcnt, 5).Value)
									case "75"
										oExcel.Cells(FacCoderowcnt, 6).Value =  "30" & chr(34)
									case "45"
										oExcel.Cells(FacCoderowcnt, 6).Value = "18" & chr(34)
									case "100"
										oExcel.Cells(FacCoderowcnt, 6).Value = "42" & chr(34)
									case "25"
										oExcel.Cells(FacCoderowcnt, 6).Value = "10" & chr(34)
									case "70"
										oExcel.Cells(FacCoderowcnt, 6).Value = "27" & chr(34)
							end Select
						end if

							Set oRSFacCodes = Nothing

							FacCoderowcnt = FacCoderowcnt + 1
						LOOP
			'END IF

	'end if
	''Jess Modifico aqui para Turkey 3/14/2017

	if ucase(trim(curCountry))="TURKEY" and not ucase(trim(curClient)) = "BEYBI" then
	''	oexcel.Range("J20").Select
	''	oexcel.Selection.EntireColumn.Insert
	''	OEXCEL.Worksheets("Sheet1").Range("J20:J21").Font.Bold = True
	''	oExcel.cells(23,10).VALUE ="IN"
	''	oExcel.cells(24,10).VALUE="TURKISH"

	''	oexcel.Range("L20").Select
	''	oexcel.Selection.EntireColumn.Insert
	''	OEXCEL.Worksheets("Sheet1").Range("L20:L21").Font.Bold = True
	''	oExcel.cells(23,12).VALUE ="IN"
	''	oExcel.cells(24,12).VALUE="TURKISH"


	''	IF OutputType = "proforma" or outputype="invoice" then
	''		datastart=23	es  23
	''	else
	''		datastart=23
	''		'datastart=25
	''	end if


	''	TotalNoRows=oExcel.ActiveSheet.UsedRange.Rows.Count
	''	rowcnt = datastart

	''	DO WHILE rowcnt <= TotalNoRows-1
	''			EnglishCurvature=oExcel.cells(rowcnt,9).VALUE


	''			qry="select * from translation  where upper(ltrim(rtrim(translation.[desc])))='" &  ucase(trim(EnglishCurvature)) & "' AND ltrim(rtrim(translation.country))='Turkey'"
	''			Set oRS = oConn.Execute(qry)
	''
	''			if not oRS.eof then
	''				oExcel.cells(rowcnt,10).VALUE = trim(oRs.Fields("trans_desc"))
	''			end if
	''
	''			EnglishCrossSec=oExcel.cells(rowcnt,11).VALUE
	''			qry="select * from translation  where upper(ltrim(rtrim(translation.[desc])))='" &  ucase(trim(EnglishCrossSec)) & "' AND ltrim(rtrim(translation.country))='Turkey'"
	''			Set oRS = oConn.Execute(qry)
	''
	''			if not oRS.eof then
	''				oExcel.cells(rowcnt,12).VALUE = trim(oRs.Fields("trans_desc"))
	''			end if
	''
	''			rowcnt = rowcnt + 1
	''	LOOP
	' hasta aqui

		'PICPOS=TotalNoRows+2

		'columnstr1=chr(64+2)
	   	'rangestr = "D"+trim(cSTR(PICPOS))
		'oExcel.Range(rangestr).Select
		'oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\graphic2.gif").Select

	end if


	if  ucase(trim(curClient)) = "BEYBI" then



		oexcel.Range("E20").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("E20:E21").Font.Bold = True
		oExcel.cells(23,5).VALUE ="IN"
		oExcel.cells(24,5).VALUE="TURKISH"


		oexcel.Range("H20").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("H20:H21").Font.Bold = True
		oExcel.cells(23,8).VALUE ="IN"
		oExcel.cells(24,8).VALUE="TURKISH"



		oexcel.Range("L20").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("L20:L21").Font.Bold = True
		oExcel.cells(23,12).VALUE ="IN"
		oExcel.cells(24,12).VALUE="TURKISH"

		oexcel.Range("N20").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("N20:N21").Font.Bold = True
		oExcel.cells(23,14).VALUE ="IN"
		oExcel.cells(24,14).VALUE="TURKISH"


		IF OutputType = "proforma" or outputype="invoice" then
			datastart=23
		else
			datastart=23
			'datastart=25
		end if


		TotalNoRows=oExcel.ActiveSheet.UsedRange.Rows.Count
		rowcnt = datastart

		DO WHILE rowcnt <= TotalNoRows-1



				EnglishType=oExcel.cells(rowcnt,4).VALUE


				qry="select * from translation  where upper(ltrim(rtrim(translation.[desc])))='" &  ucase(trim(EnglishType)) & "' AND ltrim(rtrim(translation.country))='Turkey'"
				Set oRS = oConn.Execute(qry)

				if not oRS.eof then
					oExcel.cells(rowcnt,5).VALUE = trim(oRs.Fields("trans_desc"))
				end if



				EnglishColor=oExcel.cells(rowcnt,7).VALUE


				qry="select * from translation  where upper(ltrim(rtrim(translation.[desc])))='" &  ucase(trim(EnglishColor)) & "' AND ltrim(rtrim(translation.country))='Turkey'"
				Set oRS = oConn.Execute(qry)

				if not oRS.eof then
					oExcel.cells(rowcnt,8).VALUE = trim(oRs.Fields("trans_desc"))
				end if


				EnglishCurvature=oExcel.cells(rowcnt,11).VALUE

				qry="select * from translation  where upper(ltrim(rtrim(translation.[desc])))='" &  ucase(trim(EnglishCurvature)) & "' AND ltrim(rtrim(translation.country))='Turkey'"
				Set oRS = oConn.Execute(qry)

				if not oRS.eof then
					oExcel.cells(rowcnt,12).VALUE = trim(oRs.Fields("trans_desc"))
				end if

				EnglishCrossSec=oExcel.cells(rowcnt,13).VALUE
				qry="select * from translation  where upper(ltrim(rtrim(translation.[desc])))='" &  ucase(trim(EnglishCrossSec)) & "' AND ltrim(rtrim(translation.country))='Turkey'"
				Set oRS = oConn.Execute(qry)

				if not oRS.eof then
					oExcel.cells(rowcnt,14).VALUE = trim(oRs.Fields("trans_desc"))
				end if

				rowcnt = rowcnt + 1
		LOOP

		'PICPOS=TotalNoRows+2

		'oExcel.Columns("A:R").EntireColumn.AutoFit
		'oExcel.ActiveSheet.Range("A18:R19").Interior.ColorIndex = 15


		'columnstr1=chr(64+2)
	   	'rangestr = "D"+trim(cSTR(PICPOS))
		'oExcel.Range(rangestr).Select


	end if






	if ucase(trim(curClient))="ABRA" then


		db_path =  "c:\sopdata_\pole_tran.xls"
		connectstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_path & "; Extended Properties=Excel 8.0;"
		Set oExcelConn = Server.CreateObject("ADODB.Connection")
		oExcelconn.Mode = 3
		oExcelconn.Open connectstr
		Session.CodePage = 852

		oexcel.Range("J20").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("J20:J21").Font.Bold = True
		oExcel.cells(20,10).VALUE ="IN"
		oExcel.cells(21,10).VALUE="POLISH"

		oexcel.Range("L20").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("L20:L21").Font.Bold = True
		oExcel.cells(20,12).VALUE ="IN"
		oExcel.cells(21,12).VALUE="POLISH"


		IF OutputType = "proforma" or outputype="invoice" then
			datastart=23
		else
			datastart=23
			'datastart=25
		end if






	'oExcel.Rows(trim(cSTR(datastart)))+":"+trim(cSTR(currow-1)))).EntireColumn.AutoFit



		TotalNoRows=oExcel.ActiveSheet.UsedRange.Rows.Count
		rowcnt = datastart

		DO WHILE rowcnt <= TotalNoRows-1
				EnglishCurvature=oExcel.cells(rowcnt,9).VALUE


				qry="SELECT * FROM [Sheet1$] where ucase(trim([Sheet1$].desc))='" &  ucase(trim(EnglishCurvature)) + "'"
				Set oRSExcel = oExcelconn.Execute(qry)

				if not oRSExcel.eof then
					oExcel.cells(rowcnt,10).VALUE = trim(oRSExcel.Fields("trans_desc"))
				else
					oExcel.cells(rowcnt,10).VALUE = "Not Found"
				end if


				EnglishCrossSec=oExcel.cells(rowcnt,11).VALUE
				qry="SELECT * FROM [Sheet1$] where [Sheet1$].desc='" &  ucase(trim(EnglishCrossSec)) + "'"
				Set oRSExcel = oExcelconn.Execute(qry)

				if not oRSExcel.eof then
					oExcel.cells(rowcnt,12).VALUE = trim(oRSExcel.Fields("trans_desc"))
				end if
				set oRSExcel = Nothing
				rowcnt = rowcnt + 1
		LOOP

		set oExcelconn = Nothing
		set oRSExcel = Nothing

		PICPOS=TotalNoRows+2

		'columnstr1=chr(64+2)
	   	rangestr = "D"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		rem oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\poland.gif").Select

	end if



		if ucase(trim(curCountry))="NICARAGUA" then


		db_path =  "c:\sopdata_\spain_tran.xls"
		connectstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & db_path & "; Extended Properties=Excel 8.0;"
		Set oExcelConn = Server.CreateObject("ADODB.Connection")
		oExcelconn.Mode = 3
		oExcelconn.Open connectstr
		Session.CodePage = 852

		oexcel.Range("E19").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("G18:G19").Font.Bold = True
		oExcel.cells(18,5).VALUE ="IN"
		oExcel.cells(19,5).VALUE="SPANISH"


		oexcel.Range("H19").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("G18:G19").Font.Bold = True
		oExcel.cells(18,8).VALUE ="IN"
		oExcel.cells(19,8).VALUE="SPANISH"




		oexcel.Range("L19").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("K18:K19").Font.Bold = True
		oExcel.cells(18,12).VALUE ="IN"
		oExcel.cells(19,12).VALUE="SPANISH"

		oexcel.Range("N19").Select
		oexcel.Selection.EntireColumn.Insert
		OEXCEL.Worksheets("Sheet1").Range("M18:M19").Font.Bold = True
		oExcel.cells(18,14).VALUE ="IN"
		oExcel.cells(19,14).VALUE="SPANISH"


		IF OutputType = "proforma" or outputype="invoice" then
			datastart=23
		else
			datastart=23
			'datastart=25
		end if






	'oExcel.Rows(trim(cSTR(datastart)))+":"+trim(cSTR(currow-1)))).EntireColumn.AutoFit



		TotalNoRows=oExcel.ActiveSheet.UsedRange.Rows.Count
		rowcnt = datastart

		DO WHILE rowcnt <= TotalNoRows-1

				EnglishType=oExcel.cells(rowcnt,4).VALUE


				qry="SELECT * FROM [Sheet1$] where ucase(trim([Sheet1$].desc))='" &  ucase(trim(EnglishType)) + "'"
				Set oRSExcel = oExcelconn.Execute(qry)

				if not oRSExcel.eof then
					oExcel.cells(rowcnt,5).VALUE = trim(oRSExcel.Fields("trans_desc"))
				else
					'oExcel.cells(rowcnt,5).VALUE = "Not Found"
				end if



				EnglishColor=oExcel.cells(rowcnt,7).VALUE


				qry="SELECT * FROM [Sheet1$] where ucase(trim([Sheet1$].desc))='" &  ucase(trim(EnglishColor)) + "'"
				Set oRSExcel = oExcelconn.Execute(qry)

				if not oRSExcel.eof then
					oExcel.cells(rowcnt,8).VALUE = trim(oRSExcel.Fields("trans_desc"))
				else
					'oExcel.cells(rowcnt,8).VALUE = "Not Found"
				end if



				EnglishCurvature=oExcel.cells(rowcnt,11).VALUE


				qry="SELECT * FROM [Sheet1$] where ucase(trim([Sheet1$].desc))='" &  ucase(trim(EnglishCurvature)) + "'"
				Set oRSExcel = oExcelconn.Execute(qry)

				if not oRSExcel.eof then
					oExcel.cells(rowcnt,12).VALUE = trim(oRSExcel.Fields("trans_desc"))
				else
					'oExcel.cells(rowcnt,12).VALUE = "Not Found"
				end if


				EnglishCrossSec=oExcel.cells(rowcnt,13).VALUE
				qry="SELECT * FROM [Sheet1$] where [Sheet1$].desc='" &  ucase(trim(EnglishCrossSec)) + "'"
				Set oRSExcel = oExcelconn.Execute(qry)

				if not oRSExcel.eof then
					oExcel.cells(rowcnt,14).VALUE = trim(oRSExcel.Fields("trans_desc"))
				end if
				set oRSExcel = Nothing
				rowcnt = rowcnt + 1
		LOOP

		set oExcelconn = Nothing
		set oRSExcel = Nothing

		PICPOS=TotalNoRows+2

		'columnstr1=chr(64+2)
		oExcel.Columns("O:Q").Select
		oExcel.Columns("O:Q").EntireColumn.AutoFit
		oExcel.ActiveSheet.Range("N18:Q19").Interior.ColorIndex = 15

	end if

	if OutputType = "factory" then

		if left(curpricetype,1)="U"  then
			cNewLabelJPG="C:\Inetpub\wwwroot\images\us.jpg"
		else
			cNewLabelJPG="C:\Inetpub\wwwroot\images\international.jpg"
		end if

		picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 1

		rangestr = "B"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert(trim(cNewLabelJPG)).Select
	end if




	qry="select Factory_Jpg,proforma_jpg from clients where upper(ltrim(rtrim(client)))='" & ucase(trim(curClient))  &"'"
	Set oJPGRS = oConn.Execute(qry)
	if not oJPGRS.eof then
		curFactoryJpg = oJPGRS.Fields("Factory_Jpg")
		curproformaJpg = oJPGRS.Fields("proforma_jpg")
	end if





	if LEN(trim(curfACTORYJPG)) > 0  and  OutputType = "factory" then

			picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 1

		rangestr = "D"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		'oExcel.ActiveSheet.Pictures.Insert(trim(curproformaJpg)).Select


	end if



	if LEN(trim(curproformaJpg)) > 0  and OutputType = "proforma" then

		picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 1

		rangestr = "D"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		'oExcel.ActiveSheet.Pictures.Insert(trim(curproformaJpg)).Select



	end if

		set oJPGRS = Nothing



	if ucase(trim(curCountry))="EGYPT" and ucase(trim(curClient)) = "IM TRADING & MARKETING" then

			picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 1

		rangestr = "D"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\egypt.jpg").Select


	end if

	if ucase(trim(curClient))="ORAN Y ASOCIADOS" then

			picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 1

		rangestr = "H"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\MexicoOran.jpg").Select


	end if

	if ucase(trim(curClient))="QRSHC" then

			picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 1

		rangestr = "H"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\QRSHC.jpg").Select


	end if



	if ucase(trim(curClient))="NICOLAS CARRILLO" then

			picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 2

		rangestr = "D"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\vetuselabel.jpg").Select


	end if

	if ucase(trim(curCountry))="INDIA"  and IndiaspoolFlag = "Y" then

		picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 2

		rangestr = "D"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\SilkSpoolsIndia.jpg").Select


	'india combined.jpg
	end if
	if instr(1,ucase(trim(curClient)),"SOLEX") > 0   then

		picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 2

		if 	IndiaspoolFlag = "Y" then
			rangestr = "A"+trim(cSTR(PICPOS))
			oExcel.Range(rangestr).Select
		else
			rangestr = "C"+trim(cSTR(PICPOS))
			oExcel.Range(rangestr).Select
		end if


		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\india label.jpg").Select

	'india combined.jpg
	end if



	if ucase(trim(curClient))="INSUMEDICAL LTDA" then

			picpos=oExcel.ActiveSheet.UsedRange.Rows.Count + 2

		rangestr = "C"+trim(cSTR(PICPOS))
		oExcel.Range(rangestr).Select
		oExcel.ActiveSheet.Pictures.Insert("C:\Inetpub\wwwroot\images\insumedica.gif").Select


	end if


	 if OutputType = "proforma" then

		if ucase(trim(curCountry))="TURKEY" or ucase(trim(curClient))="ABRA" then

			oExcel.ActiveSheet.Range("I"+ trim(cSTR(currow))+":O"+trim(cSTR(currow))).Interior.ColorIndex = 6

		end if

	 End if
'		else
'			oExcel.ActiveSheet.Range("I"+ trim(cSTR(currow))+":M"+trim(cSTR(currow))).Interior.ColorIndex = 6
'		end if
'
'			'oExcel.Selection.Font.Bold = .t.
'	end if



		rangestr = "A1:D" & cstr(currow)
		oExcel.Range(rangestr).Select
		With oExcel.Selection.Font
		   .Name = "Arial"
		   '.Size = 18
		End With


	if OutputType = "proforma" then
		lastrow = oExcel.ActiveSheet.UsedRange.Rows.Count
		With oExcel.ActiveSheet.Range(oExcel.ActiveSheet.Cells(2, 1), oExcel.ActiveSheet.Cells(lastrow, 26))
			   .Font.Size = 18

		End With
		oExcel.Worksheets("Sheet1").Range("D10").Font.Size = 24
		if not left(passedPO,1)="Q" then
			OEXCEL.Worksheets("Sheet1").Range("E10").Interior.ColorIndex = 8
		end if
		OEXCEL.Worksheets("Sheet1").Range("D10").Interior.Pattern = xlSolid
		oExcel.ActiveSheet.Range("A22").HorizontalAlignment = xlLeft
		' oExcel.ActiveSheet.Range("D13").select
			' With oExcel.Selection.Borders(xlEdgeLeft)
						' .LineStyle = xlNone

						' .ColorIndex = xlAutomatic
			' End With
			' With oExcel.Selection.Borders(xlEdgeRight)
						' .LineStyle = xlNone

						' .ColorIndex = xlAutomatic
			' End With

			' oExcel.ActiveSheet.Range("C13").select
			' With oExcel.Selection.Borders(xlEdgeRight)
						' .LineStyle = xlNone

						' .ColorIndex = xlAutomatic
			' End With
			' With oExcel.Selection.Borders(xlEdgeLeft)
						' .LineStyle = xlNone

						' .ColorIndex = xlAutomatic
			' End With

		' oExcel.ActiveSheet.Range("D14").select
		' With oExcel.Selection.Borders(xlEdgeLeft)
						' .LineStyle = xlNone

						' .ColorIndex = xlAutomatic
		' End With
		' With oExcel.Selection.Borders(xlEdgeRight)
						' .LineStyle = xlNone

						' .ColorIndex = xlAutomatic
		' End With

		oExcel.ActiveSheet.Range("B13").select
		With oExcel.Selection.Borders(xlEdgeRight)
						.LineStyle = xlNone

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("B14").select
		With oExcel.Selection.Borders(xlEdgeRight)
					.LineStyle = xlNone

					.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("E14").select
		With oExcel.Selection.Borders(xlEdgeRight)
						.LineStyle = xlNone

						.ColorIndex = xlAutomatic
		End With
oExcel.ActiveSheet.Range("E13").select
			With oExcel.Selection.Borders(xlEdgeRight)
						.LineStyle = xlNone

						.ColorIndex = xlAutomatic
		End With
	end if



	'do not modify spreadsheet beyond this point


	if  not oExcel.ActiveSheet.Range("N22").value="NOTES" then
		oExcel.Columns("B:K").EntireColumn.AutoFit
	else
		oExcel.Columns("B:K").EntireColumn.AutoFit
	end if



	if not disp_mess = "" then
		oExcel.cells(8,4).value = trim(disp_mess)

		OEXCEL.Worksheets("Sheet1").Range("D8:D8").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("D8:D8").Font.Bold = True
		OEXCEL.Worksheets("Sheet1").Range("D8:D8").HorizontalAlignment = xlCenter
		oExcel.Worksheets("Sheet1").Range("D8:J8").MergeCells = True

	end if



	if ucase(trim(curCountry))="EGYPT" and ucase(trim(curClient)) = "IM TRADING & MARKETING" then
		oExcel.cells(9,3).value = "** USE EGYPTIAN ART WORK **"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if



	if ucase(trim(curClient))="HEALTHCARE PRODUCTS CENTROAMERICA, S. DE R.L." then
		if  OutputType = "factory" or OutputType="factorycomp"  or OutputType="factorycomp2"  then
			oExcel.cells(9,3).value = "** USE NEW PACKAGING **"
		else
			oExcel.cells(9,5).value = "** USE NEW PACKAGING **"
		end if
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if




	if ucase(trim(curCountry))="TURKEY" then
		'JESS modifico aqui para Turkey
	''	oExcel.cells(9,3).value = "** USE TURKISH ART WORK **"
	''	oExcel.Range("C3").Select
	''	OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
	''	OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if

	if ucase(trim(curClient))="HELAL TEB" then
		oExcel.cells(9,3).value = "**USE MADE IN USA BOXES**"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if

	if ucase(trim(curClient))="ORAN Y ASOCIADOS" then
		oExcel.cells(14,6).value = "**USE MARBETES**"
		oExcel.Range("F6").Select
		OEXCEL.Worksheets("Sheet1").Range("F14:F14").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("F14:F14").Font.Bold = True
	end if

	if ucase(trim(curClient))="QRSHC" then
		oExcel.cells(14,6).value = "**USE BRAND NAME LABELS**"
		oExcel.Range("F6").Select
		OEXCEL.Worksheets("Sheet1").Range("F14:F14").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("F14:F14").Font.Bold = True
	end if

	if ucase(trim(curClient))="ABRA" then
		oExcel.cells(9,3).value = "** USE POLISH ART WORK AND INSERTS**"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if


	if ucase(trim(curCountry))="INDIA" and IndiaspoolFlag = "Y" then
		oExcel.cells(9,3).value = "** USE INDIAN ART WORK **"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if

	if instr(1,ucase(trim(curClient)),"SOLEX") > 0   then
		oExcel.cells(9,3).value = "** USE SOLEX ART WORK **"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if




	if ucase(trim(curClient))="NICOLAS CARRILLO" then
		oExcel.cells(9,3).value = "** USE NICOLAS CARRILLO ART WORK **"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if


	if ucase(trim(curClient))="INSUMEDICAL LTDA" then
		oExcel.cells(9,3).value = "** USE INSUMEDICAL ART WORK **"
		oExcel.Range("C3").Select
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.SIZE = 32
		OEXCEL.Worksheets("Sheet1").Range("C9:C9").Font.Bold = True
	end if




	if OutPutType = "packing" then
		oexcel.Range("M19").Select
		oexcel.Selection.EntireColumn.Delete
	end if

'oExcel.cells(10,10).VALUE = ucase(trim(curClient))


	if OutPutType = "factory" then
		'if ucase(trim(curCountry))="TURKEY" or ucase(trim(curClient))="ABRA" or ucase(trim(curCountry))="USA" THEN
		if  ucase(trim(curClient))="ABRA" or ucase(trim(curCountry))="USA" THEN

			oexcel.Range("N19").Select
			oexcel.Selection.EntireColumn.Delete
			oexcel.Range("N19").Select
			oexcel.Selection.EntireColumn.Delete


		ELSE

			if ucase(trim(curClient))="BEYBI" then

				oexcel.Range("N20").Select
				oexcel.Selection.EntireColumn.Delete
				oexcel.Range("O20").Select
				oexcel.Selection.EntireColumn.Delete

				oexcel.Range("O20").Select
				oexcel.Selection.EntireColumn.Delete


			else

				oexcel.Range("L20").Select
				oexcel.Selection.EntireColumn.Delete
				oexcel.Range("L20").Select
				oexcel.Selection.EntireColumn.Delete
			end if
		END IF
	end if



	'oexcel.ActiveWindow.View = xlPageBreakPreview
	if ucase(trim(curCountry))="TURKEY" or ucase(trim(curClient))="ABRA" then
		oexcel.ActiveWindow.View = xlPageBreakPreview
		oExcel.ActiveWindow.Zoom = 50
		oExcel.ActiveSheet.PageSetup.zoom = 50

		if not OutputType = "invoice" then
			oExcel.ActiveSheet.Range("A"+ trim(cSTR(datastart-2))+":O"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15
		end if


		oExcel.Worksheets("Sheet1").Range("A"+ trim(cSTR(datastart-2))+":O"+trim(cSTR(datastart-1))).Font.Bold = True
		oExcel.Worksheets("Sheet1").Range("A18:O19").HorizontalAlignment = xlCenter

	else
			if  OutputType = "factory" or OutputType="factorycomp"  or OutputType="factorycomp2"   then

				if instr(1,ucase(trim(curClient)),"SOLEX") = 0 and not ucase(trim(curCountry))="USA" then
					oexcel.Columns("B:B").Select
					'oexcel.Selection.delete Shift:=xlToLeft
					oexcel.Selection.delete
					OEXCEL.Worksheets("Sheet1").Range("P18:P19").Font.Bold = True
				end if
			end if


		if  OutputType = "proforma"  and not left(passedPO,1)="Q" and not left(curpricetype,1)="U" then
					'oexcel.Columns("B:B").Select
					'oexcel.Selection.delete
		end if




		oexcel.ActiveWindow.View = xlPageBreakPreview


		if OutputType = "proforma" then
			oExcel.ActiveWindow.Zoom = 50
			oExcel.ActiveSheet.PageSetup.zoom = 46
		else

			oExcel.ActiveWindow.Zoom = 75
			if OutputType="invoice" or OutputType="factory" or OutputType="factorycomp" or OutputType="factorycomp2" then
				oExcel.ActiveSheet.PageSetup.zoom = 100
			else

				oExcel.ActiveSheet.PageSetup.zoom = 75
			end if
		end if
		if not OutputType = "invoice" and not OutputType = "factory"  and not  OutputType="factorycomp" and not OutputType="factorycomp2" then
			oExcel.ActiveSheet.Range("A"+ trim(cSTR(datastart-2))+":M"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15
		end if

		if  OutputType = "factory"   then
			'oExcel.ActiveSheet.Range("A"+ trim(cSTR(datastart-2))+":N"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15
			oExcel.ActiveSheet.Range("A"+ trim(cSTR(datastart-2))+":M"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15
		end if

		if  OutputType = "factorycomp" or OutputType = "factorycomp2"  then
			'oExcel.ActiveSheet.Range("A"+ trim(cSTR(datastart-2))+":N"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15
			'oExcel.ActiveSheet.Range("A"+ trim(cSTR(datastart-2))+":P"+trim(cSTR(datastart-1))).Interior.ColorIndex = 15


		end if

		oExcel.Worksheets("Sheet1").Range("A"+ trim(cSTR(datastart-2))+":O"+trim(cSTR(datastart-1))).Font.Bold = True
		oExcel.Worksheets("Sheet1").Range("A"+ trim(cSTR(datastart-2))+":O"+trim(cSTR(datastart-1))).HorizontalAlignment = xlCenter

	end if
	 'Set oExcel.ActiveSheet.VPageBreaks(1).Location = oExcel.Range("N1")
			oExcel.ActiveSheet.Range("C13:I13").Select
					With oExcel.Selection.Borders(xlEdgeLeft)
						.LineStyle = xlnone

						.ColorIndex = xlAutomatic
					End With
					With oExcel.Selection.Borders(xlEdgeTop)
						.LineStyle = xlnone

						.ColorIndex = xlAutomatic
					End With
					With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlnone

						.ColorIndex = xlAutomatic
					End With
					With oExcel.Selection.Borders(xlEdgeRight)
						.LineStyle = xlnone

						.ColorIndex = xlAutomatic
					End With




	oExcel.ActiveWorkbook.SaveAs xls_path
	oExcel.ActiveWorkbook.Close(0)
	oExcel.Quit()
	'This really should be the path created up herre but it will not redirect to an absoulte path like c:inetpu\wroot\exceldump
	redirectpath="exceldump/" & tempfile
	response.write "<html>"
	response.write "<body onload=""document.location.href='" & redirectpath & "'"">"
	response.write "</html>"






	oconn.close


	set oExcel = nothing
	Set oRs = nothing
	Set oConn = nothing

sub doFacdiscount()

currow = currow + 1
oexcel.Range("L"+TRIM(cstr(Currow))).Select
oexcel.Selection.EntireRow.Insert
oexcel.Range("I"+TRIM(cstr(Currow))).value = "BULK PACKAGE (-$0.15/dz)"
discval =    pototqty * 0.15


oexcel.Range("O"+TRIM(cstr(Currow))).value = 0-discVal
rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
oexcel.Range(rangestr).Select
With oexcel.Selection
	 .HorizontalAlignment = xlCenter
End With
'marker

rangestr = "I"+trim(cSTR(currow))+":P"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True


		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With

    With oexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With
    'With oexcel.Selection.Borders(xlInsideHorizontal)
    ' '   .LineStyle = 1
       ' .Weight = xlThin
     '   '.ColorIndex = xlAutomatic
    'End With


end sub







sub dodiscount()

currow = currow + 1
oexcel.Range("L"+TRIM(cstr(Currow))).Select
oexcel.Selection.EntireRow.Insert
oexcel.Range("I"+TRIM(cstr(Currow))).value = "DISCOUNT" + " ("+trim(cstr(curDiscount)) + "%)"
discval =    pototvalue - (pototvalue*(100-cdbl(curdiscount)))/100
oexcel.Range("M"+TRIM(cstr(Currow))).value = discVal
rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
oexcel.Range(rangestr).Select
With oexcel.Selection
	 .HorizontalAlignment = xlCenter
End With
'marker

rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True


		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With




end sub

sub dofreight()
currow = currow + 1
oexcel.Range("L"+TRIM(cstr(Currow))).Select
oexcel.Selection.EntireRow.Insert
oexcel.Range("I"+TRIM(cstr(Currow))).value = "Freight"
oexcel.Range("M"+TRIM(cstr(Currow))).value = formatcurrency(curfreight)
rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
oexcel.Range(rangestr).Select
With oexcel.Selection
	 .HorizontalAlignment = xlCenter
End With
rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True

		rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With




end sub


sub doadjust()
currow = currow + 1
oexcel.Range("L"+TRIM(cstr(Currow))).Select
response.write(oexcel.Range("L"+TRIM(cstr(Currow))))
oexcel.Selection.EntireRow.Insert
if not trim(curAdjustDesc)="" then
	oexcel.Range("I"+TRIM(cstr(Currow))).value = trim(curAdjustDesc)
else
	if cdbl(curadjust) < 0 then
		oexcel.Range("I"+TRIM(cstr(Currow))).value = "Discount"
		else
		oexcel.Range("I"+TRIM(cstr(Currow))).value = "Adjustment"
	end if
end if

oexcel.Range("M"+TRIM(cstr(Currow))).value = formatcurrency(curadjust)
rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))

oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
oexcel.Range(rangestr).Select
With oexcel.Selection
	 .HorizontalAlignment = xlCenter
End With
rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True

		rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With



end sub




sub donetfactotal()
currow = currow + 1
oexcel.Range("L"+TRIM(cstr(Currow))).Select
oexcel.Selection.EntireRow.Insert
oexcel.Range("I"+TRIM(cstr(Currow))).value = "Total"


net_total  = pototcost - discval


response.write net_total



'FormatCurrency("0" & TotPersonnel)
net_total = FormatCurrency(net_total,0)




oexcel.Range("O"+TRIM(cstr(Currow))).value = net_total
rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
oexcel.Range(rangestr).Select
With oexcel.Selection
	 .HorizontalAlignment = xlCenter
End With
rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True

		rangestr = "I"+trim(cSTR(currow))+":P"+trim(cSTR(currow))
		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
    With oexcel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
    End With



end sub













sub donettotal()
currow = currow + 1
oexcel.Range("L"+TRIM(cstr(Currow))).Select
oexcel.Selection.EntireRow.Insert
oexcel.Range("I"+TRIM(cstr(Currow))).value = "TOTAL"

net_total = pototvalue
if not trim(curdiscount)="" then
	discval =    pototvalue - (pototvalue*(100-cdbl(curdiscount)))/100
	net_total  = net_total - discval
end if

if not cdbl(curadjust)=0 then
	net_total  = net_total + cdbl(curadjust)
end if

if not cdbl(curfreight)=0 then
	net_total  = net_total + cdbl(curfreight)
end if

net_total = formatcurrency(net_total)

oexcel.Range("M"+TRIM(cstr(Currow))).value = net_total
rangestr = "I"+trim(cSTR(currow))+":J"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
oexcel.Range(rangestr).Select
With oexcel.Selection
	 .HorizontalAlignment = xlCenter
End With
rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
oExcel.Worksheets("Sheet1").Range(rangestr).Font.Bold = True

		rangestr = "I"+trim(cSTR(currow))+":M"+trim(cSTR(currow))
		oExcel.Range(rangestr).Select
		oExcel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
		oExcel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
		With oExcel.Selection.Borders(xlEdgeLeft)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeTop)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With
		With oExcel.Selection.Borders(xlEdgeRight)
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = xlAutomatic
		End With




end sub
sub Dibujar()

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).value="*Production Time: 4-12 Weeks"'+variable+" *Translation from competitor codes cannot be guaranteed"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+2-nExtraLines))).font.bold=True

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).value="*Minimum order 50 boxes per code. "'+variable+" *Translation from competitor codes cannot be guaranteed"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+3-nExtraLines))).font.bold=True

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).value=variable
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+4-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).font.bold=True
		'
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).value="*Translation from competitor codes cannot be guaranteed"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+5-nExtraLines))).font.bold=True

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).value="*By accepting this Proforma Invoice, the purchaser agrees to pay for these medical devices on the above mentioned payment terms."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).font.bold=True

		   oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).value="*Specialty Items may have longer production times."
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).font.bold=True

		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).value="on the above mentioned payment terms."
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+6-nExtraLines))).font.bold=True

		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).font.size=26
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+7-nExtraLines))).value=""

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.size=26
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).value="BANKING INFORMATION: Please send payment to:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+9-nExtraLines))).font.bold=True

		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).value="First American Bank, Elk Grove Village, IL USA"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).value="First American Bank, 2295 Galiano St. Coral Gables, FL USA"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+10-nExtraLines))).font.bold=True

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).value="With the Federal Reserve Bank"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+11-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).value="ABA no. 071-922-777"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+12-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).value="For Credit to Trade Finance Division"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+13-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).value="Account 78 11 57 1101: Referencing Demetech Corporation"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).value="Account 78 11 57 1118: Referencing Demetech Corporation"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+14-nExtraLines))).font.size=26


		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).value="AGREED & ACCEPTED:"
		'xlContinuous




		'oExcel.ActiveSheet.Range("C13:I13").Select
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+17-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("E"+trim(cSTR(currow+17-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("F"+trim(cSTR(currow+17-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+19-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+20-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+21-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+21-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).value="Buyer:                            " & trim(curClient)
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).value="Signature:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).value="Print Name:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).value="Title:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+23-nExtraLines))).value="Date:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+23-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+23-nExtraLines))).font.size=26


		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+25-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+25-nExtraLines))).value=trackstr


		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+27-nExtraLines))).value="Deposits are Non-refundable, all sales are finals, no returns, confirmed Purchase Orders can not be modified or canceled. Full payment must be received prior to shipment of order.  Interest at a rate of 1 and 1/2 Percent per month will be charged on all delinquent accounts which are not paid when due.   In addition, you will also be responsible for any legal costs incurred by DemeTech Corporation, including attorney's fees, in collecting any past due balances.Attorney's fees in the event of Dispute: In the event of any dispute between parties arising out of or related to this contract, the prevailing party in any litigation shall be entitled to recover reasonable attorney's fees"
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+27-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+27-nExtraLines))).font.size=30
		'rangestr = "I"+trim(cSTR(currow-2))+":J"+trim(cSTR(currow-2))
		'oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
		rangestr = "D"+trim(cSTR(currow+27-nExtraLines))+":O"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		 With oExcel.Selection
			.HorizontalAlignment = xlLeft
			.VerticalAlignment = xlTop
			.RowHeight = 150
			.WrapText = True
			.MergeCells = True
	    End With
		rangestr = "C"+trim(cSTR(currow+27-nExtraLines))+":C"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		With oExcel.Selection
				.Delete(-4159)
		end With
		if cdbl(curadjust)=0 and cdbl(curFreight)=0  then
		rangestr = "B"+trim(cSTR(currow+27-nExtraLines))+":B"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		With oExcel.Selection
				.Delete(-4159)
		end With
		rangestr = "A"+trim(cSTR(currow+27-nExtraLines))+":A"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		With oExcel.Selection
				.Delete(-4159)
		end With
		end if
		  'Selection.EntireColumn.Hidden = True
		if  OutputType = "proforma"  and not left(passedPO,1)="Q" and not left(curpricetype,1)="U" then
					' oExcel.Columns("E:E").Select
					' With oExcel.Selection
						' .EntireColumn.Hidden = True
					' end  With
					' oExcel.Columns("E:E").Select
					' oexcel.Selection.delete
					rangestr = "B1:B"+trim(cSTR(currow+27-1))
					response.Write(rangestr)
					'oExcel.ActiveSheet.Range(rangestr).Select
					'With oExcel.Selection
					'.EntireColumn.Delete
		'end With
		end if
		'Excel.ActiveSheet.Range("A"+trim(cSTR(currow+25-nExtraLines)):"O"+trim(cSTR(currow+25-nExtraLines))).MergeCells = True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+26-nExtraLines))).value="Full payment must be received prior to shipment of order.  Interest at a rate of 1 and 1/2 Percent per month"
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+26-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+26-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+27-nExtraLines))).value="will be charged on all delinquent accounts which are not paid when due.   In addition, you will also be responsible for "
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+27-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+27-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+28-nExtraLines))).value="any legal costs incurred by DemeTech Corporation, including attorney's fees, in collecting any past due balances."
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+28-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+28-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+29-nExtraLines))).value="Attorney's fees in the event of Dispute: In the event of any dispute between parties arising out of or related to this "
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+29-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+29-nExtraLines))).font.size=30
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+30-nExtraLines))).value="contract, the prevailing party in any litigation shall be entitled to recover reasonable attorney's fees "
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+30-nExtraLines))).font.bold=True
		' oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+30-nExtraLines))).font.size=30
end sub
sub Dibujar2()
oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+16-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+17-nExtraLines))).value="AGREED & ACCEPTED:"
		'xlContinuousd




		'oExcel.ActiveSheet.Range("C13:I13").Select
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+18-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("E"+trim(cSTR(currow+18-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("F"+trim(cSTR(currow+18-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+20-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+21-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+22-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+23-nExtraLines))).select
			With oExcel.Selection.Borders(xlEdgeBottom)
						.LineStyle = xlContinuous

						.ColorIndex = xlAutomatic
		End With

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).value="Buyer:                            " & trim(curClient)
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+18-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).value="Signature:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+20-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).value="Print Name:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+21-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).value="Title:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+22-nExtraLines))).font.size=26

		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+23-nExtraLines))).value="Date:"
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+23-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+23-nExtraLines))).font.size=26


		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+25-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+25-nExtraLines))).value=trackstr


		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+27-nExtraLines))).value="Deposits are Non-refundable, all sales are finals, no returns, confirmed Purchase Orders can not be modified or canceled. Full payment must be received prior to shipment of order.  Interest at a rate of 1 and 1/2 Percent per month will be charged on all delinquent accounts which are not paid when due.   In addition, you will also be responsible for any legal costs incurred by DemeTech Corporation, including attorney's fees, in collecting any past due balances.Attorney's fees in the event of Dispute: In the event of any dispute between parties arising out of or related to this contract, the prevailing party in any litigation shall be entitled to recover reasonable attorney's fees"
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+27-nExtraLines))).font.bold=True
		oExcel.ActiveSheet.Range("D"+trim(cSTR(currow+27-nExtraLines))).font.size=30
		'rangestr = "I"+trim(cSTR(currow-2))+":J"+trim(cSTR(currow-2))
		'oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
		rangestr = "D"+trim(cSTR(currow+27-nExtraLines))+":O"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		 With oExcel.Selection
			.HorizontalAlignment = xlLeft
			.VerticalAlignment = xlTop
			.RowHeight = 150
			.WrapText = True
			.MergeCells = True
	    End With
		rangestr = "C"+trim(cSTR(currow+27-nExtraLines))+":C"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		With oExcel.Selection
				.Delete(-4159)
		end With
		if cdbl(curadjust)=0 and cdbl(curFreight)=0  then
		rangestr = "B"+trim(cSTR(currow+27-nExtraLines))+":B"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		With oExcel.Selection
				.Delete(-4159)
		end With
		rangestr = "A"+trim(cSTR(currow+27-nExtraLines))+":A"+trim(cSTR(currow+27-nExtraLines))
		oExcel.ActiveSheet.Range(rangestr).Select
		With oExcel.Selection
				.Delete(-4159)
		end With
		end if
		  'Selection.EntireColumn.Hidden = True
		if  OutputType = "proforma"  and not left(passedPO,1)="Q" and not left(curpricetype,1)="U" then
					' oExcel.Columns("E:E").Select
					' With oExcel.Selection
						' .EntireColumn.Hidden = True
					' end  With
					' oExcel.Columns("E:E").Select
					' oexcel.Selection.delete
					rangestr = "B1:B"+trim(cSTR(currow+27-1))
					response.Write(rangestr)
					'oExcel.ActiveSheet.Range(rangestr).Select
					'With oExcel.Selection
					'.EntireColumn.Delete
		'end With
		end if
				' response.Write(rangestr)
			'oExcel.ActiveSheet.Range("B:B").EntireColumn.Hidden = True
		'oExcel.Worksheets("Sheet1").Range(rangestr).MergeCells = True
		'oExcel.Worksheets("Sheet1").Range("A"+trim(cSTR(currow+26-nExtraLines)):"O"+trim(cSTR(currow+26-nExtraLines))).MergeCells = True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+27-nExtraLines))).value="Full payment must be received prior to shipment of order.  Interest at a rate of 1 and 1/2 Percent per month"
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+27-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+27-nExtraLines))).font.size=30
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+28-nExtraLines))).value="will be charged on all delinquent accounts which are not paid when due.   In addition, you will also be responsible for "
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+28-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+28-nExtraLines))).font.size=30
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+29-nExtraLines))).value="any legal costs incurred by DemeTech Corporation, including attorney's fees, in collecting any past due balances."
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+29-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+29-nExtraLines))).font.size=30
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+30-nExtraLines))).value="Attorney's fees in the event of Dispute: In the event of any dispute between parties arising out of or related to this "
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+30-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+30-nExtraLines))).font.size=30
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+31-nExtraLines))).value="contract, the prevailing party in any litigation shall be entitled to recover reasonable attorney's fees "
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+31-nExtraLines))).font.bold=True
		'oExcel.ActiveSheet.Range("A"+trim(cSTR(currow+31-nExtraLines))).font.size=30
end sub
function GenBarCodeNum(demecode)
	'how to call it
 	'oSheet.Cells(rowcnt, 27).Value	 = "*"+GenBarCodeNum(oSheet.Cells(rowcnt, 3).Value)+"*"
	qry="select prodcode from livedata where code ='" + trim(demecode) +"'"
	Set ProdCodeRS = oConn.Execute(qry)

	if not ProdCodeRS.eof then
		GenBarCodeNum=gencheckdigit("0652927"+cStr(ProdCodeRS.fields("prodcode").value))
	end if
	'Set ProdCodeRS = Nothing
end function


function gencheckdigit(barcode)

        nNumLen = Len(Trim(barcode))
        'do the even numbers
        curpos = nNumLen
        eventot = 0
        Do While curpos > 0
            curdigit = Mid(barcode, curpos, 1)
			curdigit =cInt(curDigit)
            eventot = eventot + curdigit
            curpos = curpos - 2
      Loop
        eventot = eventot * 3

        'do the oddnumbers
        curpos = nNumLen - 1
        oddtot = 0
        Do While curpos > 0
            curdigit = CInt(Mid(barcode, curpos, 1))
            oddtot = oddtot + curdigit
            curpos = curpos - 2

        Loop
        nTotal = eventot + oddtot
        'get the next has mutiple of 10
		nRemainder = nTotal Mod 10

		if nRemainder = 0 then
			gencheckdigit = barcode + "0"
		else
	        gencheckdigit = +barcode + cstr(10 - nRemainder)
		end if
end function

'Function to check the needle code for the product code and update

function selectneedle(demetechcode)
'response.Write("In selectneedle function"&"<br/>")
presentcode=demetechcode
query="select Needlecode from Orders where Productcode='" & presentcode &"'"
set presentquery=oConn.Execute(query)
if not presentquery.eof then
oExcel.ActiveSheet.Range("N"+trim(cSTR(currow))).value=trim(presentquery.Fields("Needlecode"))
else
oExcel.ActiveSheet.Range("N"+trim(cSTR(currow))).value="Not Available"
end if
end function


%>
