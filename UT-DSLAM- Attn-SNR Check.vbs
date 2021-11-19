# $language = "VBScript"
# $interface = "1.0"

'******************************     UT DSLAM Port Attenuation-SNR Check Script                                   ************************


  Const ForReading = 1
  Const ForWriting = 2

Sub main
  crt.screen.synchronous = true

  Dim app, wb, wb2, ws, ws2, row, nxtrow, utrow, mathalu1, mathalu2, zkram, nextzkram, ukram, nextukram, utslot(500), utport(500)
	nextzkram = 1
	nextukram = 1
  	nxtrow = 3
	utrow = 3

	Set app = CreateObject("Excel.Application")
	Set wb = app.Workbooks.Add
	Set ws = wb.Worksheets(1)
	Set wb2 = app.Workbooks.Add
	Set ws2 = wb2.Worksheets(1)
                '---------------------------------------------------------- Sheet Headings ----------------------------------------------
	ws.Cells(1, 2).Value = ">>>>>"
        	ws.Cells(1, 3).Value = " Script "
	ws.Cells(1, 4).Value = "Prepared By :"
	ws.Cells(1, 5).Value = "Rajiv "
	ws.Cells(1, 6).Value = "Kubavat"
	ws.Cells(1, 7).Value = "~~~~~"
	ws.Cells(1, 8).Value = "( TTA - "
	ws.Cells(1, 9).Value = "Gujarat"
	ws.Cells(1, 10).Value = "Circle"
	ws.Cells(1, 11).Value = "Broadband"
	ws.Cells(1, 12).Value = "- NOC )"

	ws.Cells(2, 1).Value = "Sr. No."
	ws.Cells(2, 2).Value = "DATE "
        	ws.Cells(2, 3).Value = "TIME "
	ws.Cells(2, 4).Value = "IP ADDRESS "
	ws.Cells(2, 5).Value = "DSLAM LABEL "
	ws.Cells(2, 6).Value = "Port No."
	ws.Cells(2, 7).Value = "Oper. Status"   
	ws.Cells(2, 8).Value = "ATU-C_SNR"
	ws.Cells(2, 9).Value = "ATU-R_SNR"
	ws.Cells(2, 10).Value = "ATU-C_Attenu."
	ws.Cells(2, 11).Value = "ATU-R_Attenu"
	ws.Cells(2, 12).Value = "ATU-C_Attain Rate"
	ws.Cells(2, 13).Value = "ATU-R_Attain Rate"
	ws.Cells(2, 14).Value = "ATU-C_Curr.Rate"
	ws.Cells(2, 15).Value = "ATU-R_Curr.Rate"
	

	'------------------------------------------------- Enter to BNG --------------------------------------------------------------------------

	crt.Session.Connect ("/telnet" & " "  & "10.226.144.1"  & " " & 23)	'--------BDR Ring-2 BNG
                'crt.Session.Connect ("/telnet" & " "  & "10.234.232.1"  & " " & 23)	'--------Mehsana Ring BNG
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
  	Set file = fso.OpenTextFile("H:\R-Tel  Tester\Input-UT-DSLAM-Attn-SNR-Check.txt", ForReading, False)
                'Set file = fso.OpenTextFile("C:\R-Tel  Tester\Input-UT-DSLAM-Attn-SNR-Check.txt", ForReading, False)
	'Set file = fso.OpenTextFile("D:\Raj-Tel Tester\Input-UT-DSLAM-Attn-SNR-Check.txt", ForReading, False)

  
 Do While file.AtEndOfStream <> True
           str = file.Readline
           crt.Screen.Send "telnet" & " " & str & vbcr
 
          row=nxtrow
              
        If crt.Screen.waitforstring ("ogin:", 10 ) Then  '------------------If DSLAM is Down / No Telnet then wait for 9 Seconds ------------------------
           	crt.Screen.Send "admin" & vbcr
           	crt.Screen.waitforstring ("ssword:")
           	crt.Screen.Send vbcr
           	crt.Screen.waitforstring ("#") 

	Dim screenrow, readline, chakra1, chakra2, ajay1, ajay2, looper1, looper2, looper3, nakal1, nakal2, asli1, asli2, portkram, sodh1
		zkram = nextzkram
		ws.Cells(row, 1).Value =  zkram
		ws.Cells(row, 2).Value =  Date
		ws.Cells(row, 3).Value =  Time
      		ws.Cells(row, 4).Value = str

		screenrow = crt.screen.Currentrow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 35 )
		ws.Cells(row, 5).Value = readline

  		crt.Screen.Send "show slot" & vbcr
  		crt.Screen.waitforstring ("#")
		
		Dim slotok, kram1, kram2, kram3, kramok, chakkar1

		ajay2 = 2
		looper1 = 0
		kram2 = 0
		kram3 = 0

		Do
			screenrow = crt.screen.Currentrow - ajay2
			readline = crt.Screen.Get(screenrow, 1, screenrow, 1 )
			ajay1 = readline

			If ajay1 = "S" Then
				kram1 = ajay2 
				Exit Do
			Else
				If ajay1 = "A" OR ajay1 = "B" OR ajay1 = "P" then
					kram2 = ajay2
				Else
					kram3 = ajay2
				End If
			End If

			ajay2 = ajay2 + 1
		Loop

		kram1 = kram1 - 1
		'kramok = kram1 - kram2
		kram2 = kram2 + 1
		
		Dim karate1, karate2

		looper1 = 0
		karate1 = 1		

		For chakkar1 = kram1 to kram2 step -1
			screenrow = crt.screen.Currentrow - chakkar1
			readline = crt.Screen.Get(screenrow, 21, screenrow, 24 )
			nakal1 = readline
					
			If nakal1 = "PLUG" then
				readline = crt.Screen.Get(screenrow, 6, screenrow, 14 )
				sodh1 = readline

				If sodh1 = "IPADSL8A " Then
					utport(looper1) = 48
					utslot(looper1) = karate1
					
				Else
					utport(looper1) = 24
					utslot(looper1) = karate1
					
				End If

				looper1 = looper1 + 1
				karate1 = karate1 + 1
			Else
				karate1 = karate1 + 1
			End If

		Next
							
				

		asli1 = looper1 - 1

		For looper2 = 0 to asli1 step +1

			crt.Screen.Send "slot " & utslot(looper2)  & vbcr
			crt.Screen.waitforstring ("#")
				
			crt.Screen.Send "port" & vbcr
			crt.Screen.waitforstring ("#")
			
			For looper3 = 1 to utport(looper2) step +1

				crt.Screen.Send "show line dsl " & looper3 & vbcr
				
        				If crt.Screen.waitforstring ("#", 2) Then

					Dim swami1

					screenrow = crt.screen.Currentrow - 9
					readline = crt.Screen.Get(screenrow, 1, screenrow, 4 )
					swami1 = readline

					If swami1 = "Port" Then
						readline = crt.Screen.Get(screenrow, 35, screenrow, 36 )	'----------------- line-9 Port no. shown--------------
						ws.Cells(row, 6).Value = "Port " & utslot(looper2) & "/" & readline
					Else
						screenrow = crt.screen.Currentrow - 7
						readline = crt.Screen.Get(screenrow, 35, screenrow, 36 )	'----------------- line-7 Port no. shown--------------
						ws.Cells(row, 6).Value = "Port " & utslot(looper2) & "/" & readline
					End If
					
					screenrow = crt.screen.Currentrow - 4
					readline = crt.Screen.Get(screenrow, 31, screenrow, 34 )	'----------------- Oper. status ------------------
					ws.Cells(row, 7).Value = readline

				Else
					crt.Screen.Send vbcr
				
					crt.Screen.waitforstring ("#")

					Dim swami2, swami3, swami4, swami5, swami6

					screenrow = crt.screen.Currentrow - 32
					readline = crt.Screen.Get(screenrow, 1, screenrow, 4 )
					swami2 = readline

					If swami2 = "Port" Then
						readline = crt.Screen.Get(screenrow, 35, screenrow, 36 )	'----------------- line-32 Port no. shown--------------
						ws.Cells(row, 6).Value = "Port " & utslot(looper2) & "/" & readline
					Else
						screenrow = crt.screen.Currentrow - 28
						readline = crt.Screen.Get(screenrow, 35, screenrow, 36 )	'----------------- line-28 Port no. shown--------------
						ws.Cells(row, 6).Value = "Port " & utslot(looper2) & "/" & readline
					End If
					
					For swami3 = 27 to 25 step-1
						screenrow = crt.screen.Currentrow - swami3
						readline = crt.Screen.Get(screenrow, 1, screenrow, 5 )
						If readline = "Opera" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	
							ws.Cells(row, 7).Value = readline
						End If
					Next
					
					For swami5 = 22 to 9 step-1

						screenrow = crt.screen.Currentrow - swami5
						readline = crt.Screen.Get(screenrow, 5, screenrow, 10 )	'--------------- check ----------------
						swami6 = readline

						If swami6 = "C SNR " Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-C SNR ----------------
							ws.Cells(row, 8).Value = readline
						Else If swami6 = "R SNR " Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-R SNR ----------------
							ws.Cells(row, 9).Value = readline
						Else If swami6 = "C Atte" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-C Attenu. ----------------
							ws.Cells(row, 10).Value = readline
						Else If swami6 = "R Atte" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-R Attenu. ----------------
							ws.Cells(row, 11).Value = readline
						Else If swami6 = "C Atta" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-C Attain rate. ----------------
							ws.Cells(row, 12).Value = readline
						Else If swami6 = "R Atta" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-R Attain rate. ----------------
							ws.Cells(row, 13).Value = readline
						Else If swami6 = "C Curr" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-C Curr. rate. ----------------
							ws.Cells(row, 14).Value = readline
						Else If swami6 = "R Curr" Then
							readline = crt.Screen.Get(screenrow, 31, screenrow, 35 )	'--------------- ATU-R Curr Rate ----------------
							ws.Cells(row, 15).Value = readline
						End If
						End If
						End If
						End If
						End If
						End If
						End If
						End If
					Next
					
				End If

				row = row + 1
			Next
			
			crt.Screen.Send "exit" & vbcr
			crt.Screen.waitforstring ("#")
			crt.Screen.Send "exit" & vbcr
			crt.Screen.waitforstring ("#")
		Next			
					
		

	                                
		'crt.Screen.waitforstring ("#")

		For looper2 = asli1 to 0 step -1
			utslot(looper2) = 0
			utport(looper2) = 0
		Next

		crt.Screen.Send "quit" & vbcr
		
		nxtrow = row
		'zkram = zkram + 1
		'nextzkram = zkram


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
	  	crt.Screen.Send chr(3) & vbcr    '----------- Sends (Ctrl + C)  to abort the try -----------------------
	 	nxtrow = row
		zkram = zkram + 1
		nextzkram = zkram

    End If

   crt.Screen.WaitForString "bdr-02>"		'------------------BDR Ring-2 BNG
   'crt.Screen.WaitForString "msn-01>"		'------------------Mehsana Ring BNG
   'crt.Screen.WaitForString "mnx-01>"		'------------------Surat BNG
   'crt.Screen.WaitForString "krx-01>"		'------------------Rajkot Ring-1 BNG

loop

    '----------------------------------------------------------------  Loop - Now Test second IP in Input List -----------------------

    wb.SaveAs("H:\R-Tel  Tester\Raj-UT-DSLAM-ATTN-SNR-result.xls")
   'wb.SaveAs("C:\R-Tel  Tester\Raj-UT-DSLAM-ATTN-SNR-result.xls")
    

    'wb.SaveAs("D:\Raj-Tel Tester\Raj-UT-DSLAM-ATTN-SNR-Result.xls")
    'wb2.SaveAs("D:\Raj-Tel Tester\Raj-TEL-UT-result.xls")

    wb.Close
    'wb2.Close
    app.Quit

    Set ws = nothing
    'Set ws2 = nothing
    Set wb = nothing
    'Set wb2 = nothing
    Set app = nothing

  crt.screen.synchronous = false

End Sub

'-----------------------------------------------((((((((    ZTE - UT DSLAM Telnet Tester Script   ))))))------------------------------------------------------
'------------------------------------------------Prepared by  : Rajiv J. Kubavat  -----------------------------------------------------------------
'------------------------------------------------------------------ TTA - Gujarat Circle Broadband NOC , Ahmedabad----------------------
'------------------------------------------------------------------ Phone : 079-26403788 --------------------------------------------------------