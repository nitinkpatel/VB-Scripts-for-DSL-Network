# $language = "VBScript"
# $interface = "1.0"

'******************************     NSN DSLAM - SNR Attenuation ckeck  Script  **************************************************

  Const ForReading = 1
  Const ForWriting = 2

Sub main
  crt.screen.synchronous = true

  Dim app, wb, wb2, ws, ws2, row, nxtrow, utrow, mathalu1, mathalu2, zkram, nextzkram, ukram, nextukram
	nextzkram = 1
	nextukram = 1
  	nxtrow = 3
	utrow = 3

	Set app = CreateObject("Excel.Application")
	Set wb = app.Workbooks.Add
	Set ws = wb.Worksheets(1)
	Set wb2 = app.Workbooks.Add
	Set ws2 = wb2.Worksheets(1)
                '---------------------------------------------------------- ZTE Sheet Headings ----------------------------------------------
	ws.Cells(1, 2).Value = ">>>>"
        	ws.Cells(1, 3).Value = " Script "
	ws.Cells(1, 4).Value = "Prepared By :"
	ws.Cells(1, 5).Value = "   Rajiv  "
	ws.Cells(1, 6).Value = "Kubavat"
	ws.Cells(1, 7).Value = "   ~~~~  "
	ws.Cells(1, 8).Value = "( TTA - "
	ws.Cells(1, 9).Value = "Gujarat"
	ws.Cells(1, 10).Value = "Circle"
	ws.Cells(1, 11).Value = "BB NOC )"
	ws.Cells(1, 12).Value = "<<<<"

	ws.Cells(2, 1).Value = "Sr. No."
	ws.Cells(2, 2).Value = "DATE "
        	ws.Cells(2, 3).Value = "TIME "
	ws.Cells(2, 4).Value = "IP ADDRESS "
	ws.Cells(2, 5).Value = "DSLAM LABEL "
	ws.Cells(2, 6).Value = "Port No."
	ws.Cells(2, 7).Value = "Admin/Oper.Status"   
	ws.Cells(2, 8).Value = "Downst-Line Attn."
	ws.Cells(2, 9).Value = "Downst-Signal Attn."
	ws.Cells(2, 10).Value = "Downst-SNR."
	ws.Cells(2, 11).Value = "Upst-Line Attn."
	ws.Cells(2, 12).Value = "Upst-Sign Attn."
	ws.Cells(2, 13).Value = "Upst-SNR."

	'------------------------------------------------- Enter to BNG --------------------------------------------------------------------------

	'crt.Session.Connect ("/telnet" & " "  & "10.234.232.1"  & " " & 23)	'--------Mehsana Ring BNG
	crt.Session.Connect ("/telnet" & " "  & "10.226.144.1"  & " " & 23)	'--------BDR Ring-2 BNG
	'crt.Session.Connect ("/telnet" & " "  & "10.230.32.1"  & " " & 23)	'--------Surat BNG
	'crt.Session.Connect ("/telnet" & " "  & "10.230.16.1"  & " " & 23)	'--------Rajkot Ring-1 BNG

           	crt.Screen.waitforstring ("ogin:")
           	crt.Screen.Send "bbmp" & vbcr
           	crt.Screen.waitforstring ("ssword:")
           	crt.Screen.Send "bbmp123" & vbcr
		
  	crt.Screen.waitforstring (">")
  	crt.Screen.Send "context mgmt" & vbcr
	crt.Screen.waitforstring (">")

  Dim fso, file, str

  	Set fso = CreateObject("Scripting.FileSystemObject")
	'Set file = fso.OpenTextFile("C:\R-Tel  Tester\Input-NSN-DSLAM-snr-attn-check.txt", ForReading, False)
  	Set file = fso.OpenTextFile("H:\R-Tel  Tester\Input-NSN-DSLAM-snr-attn-check.txt", ForReading, False)
	'Set file = fso.OpenTextFile("D:\Raj-Tel Tester\Input-ZTE-UT-DSLAM.txt", ForReading, False)

  
 Do While file.AtEndOfStream <> True
           str = file.Readline
           crt.Screen.Send "telnet" & " " & str & vbcr
 
     row=nxtrow
              
               

     If crt.Screen.waitforstring ("ogin:", 9 ) Then  '------------------If DSLAM is Down / No Telnet then wait for 9 Seconds ------------------------
           crt.Screen.Send "root" & vbcr
           	
           crt.Screen.waitforstring ("ssword:")
           crt.Screen.Send "vertex25" & vbcr
           crt.Screen.waitforstring (">") 

	Dim screenrow, readline, chakra1, chakra2, ajay1, ajay2, nsnslot(50), nsncard(50), looper1, looper2, looper3, nakal1, nakal2, asli1, asli2, portkram, sodh1
		zkram = nextzkram
		ws.Cells(row, 1).Value =  zkram
		ws.Cells(row, 2).Value =  Date
		ws.Cells(row, 3).Value =  Time
      		ws.Cells(row, 4).Value = str

		crt.Screen.Send "enable" & vbcr
		crt.Screen.waitforstring ("#")
		screenrow = crt.screen.Currentrow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 35 )
		ws.Cells(row, 5).Value = readline

  		crt.Screen.Send "show slot-overview | grep unlocked" & vbcr
  		crt.Screen.waitforstring ("#")
		
		ajay2 = 1
		looper1 = 0
		Do
			screenrow = crt.screen.Currentrow - ajay2
			readline = crt.Screen.Get(screenrow, 45, screenrow, 52 )
			ajay1 = readline
			If not ajay1 = "unlocked" Then Exit Do
			readline = crt.Screen.Get(screenrow, 20, screenrow, 21 )
			nakal1 = readline
			If nakal1 = "72" then
				readline = crt.Screen.Get(screenrow, 4, screenrow, 5 )
				sodh1 = readline
				If sodh1 < 10 Then
					readline = crt.Screen.Get(screenrow, 4, screenrow, 4 )
				end If

				nsnslot(looper1) = readline	
				looper1 = looper1 + 1
			End If
			ajay2=ajay2+1
		Loop

		
		asli1 = looper1 - 1

		For looper2 = asli1 to 0 step -1

			For looper3 = 1 to 72 step + 1

				crt.Screen.Send "show lre " & nsnslot(looper2) & "/" & looper3 & " " & "xdsl  band-table" & vbcr
        				crt.Screen.waitforstring ("#")
				screenrow = crt.screen.Currentrow - 13
				readline = crt.Screen.Get(screenrow, 20, screenrow, 24 )
				ws.Cells(row, 6).Value = "Port " & readline
				screenrow = crt.screen.Currentrow - 12
				readline = crt.Screen.Get(screenrow, 20, screenrow, 35 )
				ws.Cells(row, 7).Value = readline

				screenrow = crt.screen.Currentrow - 7
				readline = crt.Screen.Get(screenrow, 7, screenrow, 10 )
				ws.Cells(row, 8).Value = readline
				readline = crt.Screen.Get(screenrow, 17, screenrow, 21 )
				ws.Cells(row, 9).Value = readline
				readline = crt.Screen.Get(screenrow, 28, screenrow, 31 )
				ws.Cells(row, 10).Value = readline

				readline = crt.Screen.Get(screenrow, 40, screenrow, 43 )
				ws.Cells(row, 11).Value = readline
				readline = crt.Screen.Get(screenrow, 50, screenrow, 53 )
				ws.Cells(row, 12).Value = readline
				readline = crt.Screen.Get(screenrow, 61, screenrow, 64 )
				ws.Cells(row, 13).Value = readline

				row = row + 1
			Next
		Next	

		crt.Screen.Send "exit" & vbcr

'---------------------------------------------------- Down DSLAM Section Start -----------------------------------------------------------
     else
		row = nxtrow
		zkram = nextzkram
		ws.Cells(row, 1).Value =  zkram
		ws.Cells(row, 2).Value =  Date
		ws.Cells(row, 3).Value =  Time
      		ws.Cells(row, 4).Value = str
	  	ws.Cells(row, 5).Value =  "DOWN"
		ws.Cells(row, 6).Value =  "DOWN"
	  	row = row + 1
	  	crt.Screen.Send chr(3) & vbcr    	'----------- Sends (Ctrl + C)  to abort the try -------------------
	 	nxtrow = row
		zkram = zkram + 1
		nextzkram = zkram

    End If

   'crt.Screen.WaitForString "msn-01>"		'------------------Mehsana Ring BNG
   crt.Screen.WaitForString "bdr-02>"		'------------------BDR Ring-2 BNG
   'crt.Screen.WaitForString "mnx-01>"		'------------------Surat BNG
   'crt.Screen.WaitForString "krx-01>"		'------------------Rajkot Ring-1 BNG

loop

    '----------------------------------------------------------------  Loop - Now Test second IP in Input List -----------------------

    wb.SaveAs("H:\R-Tel  Tester\Raj-NSN-SNR-Attenuation-result.xls")
    'wb.SaveAs("C:\R-Tel  Tester\Raj-NSN-SNR-Attenuation-result.xls")
    'wb2.SaveAs("H:\R-Tel  Tester\Raj-TEL-UT-result.xls")

    wb.Close
    wb2.Close
    app.Quit

    Set ws = nothing
    Set ws2 = nothing
    Set wb = nothing
    Set wb2 = nothing
    Set app = nothing

  crt.screen.synchronous = false

End Sub

'-----------------------------------------------((((((((    NSN DSLAM - SNR Attenuation ckeck  Script   ))))))------------------------------------------------------
----