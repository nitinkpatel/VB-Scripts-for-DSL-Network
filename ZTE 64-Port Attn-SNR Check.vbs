# $language = "VBScript"
# $interface = "1.0"

'******************************     ZTE 64 Port DSLAM Attenuation-SNR Tester  Script  ************************

  Const ForReading = 1
  Const ForWriting = 2

Sub main
  crt.screen.synchronous = true

  Dim app, wb, wb2, ws, ws2, row, nxtrow, utrow, mathalu1, mathalu2, zkram, nextzkram, ukram, nextukram, zteslot(500), ztecard(500)
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
		
	ws.Cells(2, 1).Value = "Sr. No."
	ws.Cells(2, 2).Value = "DATE "
        ws.Cells(2, 3).Value = "TIME "
	ws.Cells(2, 4).Value = "IP ADDRESS "
	ws.Cells(2, 5).Value = "DSLAM LABEL "
	ws.Cells(2, 6).Value = "Port No."
	ws.Cells(2, 7).Value = "Admin/Operation"

	ws.Cells(2, 8).Value = "ATU-C-SNR"
	ws.Cells(2, 9).Value = "ATU-R-SNR"
	ws.Cells(2, 10).Value = "ATU-C-Attenu"
	ws.Cells(2, 11).Value = "ATU-R-Attenu"

	ws.Cells(2, 12).Value = "ATU-C-SNR.-org"
	ws.Cells(2, 13).Value = "ATU-R-SNR-org"
	ws.Cells(2, 14).Value = "ATU-C-Atten-org"
	ws.Cells(2, 15).Value = "ATU-R-Atten-org"

	ws.Cells(2, 16).Value = "ATU-C-Rate."
	ws.Cells(2, 17).Value = "ATU-R-Rate"
	ws.Cells(2, 18).Value = "ATU-C-O/P Pwr."
	ws.Cells(2, 19).Value = "ATU-R-O/P Pwr"
	

	'------------------------------------------------- Enter to BNG --------------------------------------------------------------------------

	
	'crt.Session.Connect ("/telnet" & " "  & "10.230.32.1"  & " " & 23)	'--------Surat BNG
	
           	crt.Screen.waitforstring ("ogin:")
           	crt.Screen.Send "bbmp" & vbcr
           	crt.Screen.waitforstring ("ssword:")
           	crt.Screen.Send "bbmp123" & vbcr
		
  	crt.Screen.waitforstring (">")
  	crt.Screen.Send "context mgmt" & vbcr
	crt.Screen.waitforstring (">")

  Dim fso, file, str

  	Set fso = CreateObject("Scripting.FileSystemObject")
  	Set file = fso.OpenTextFile("H:\R-Tel  Tester\Input-ZTE-64p-Attn-SNR-Check.txt", ForReading, False)
	'Set file = fso.OpenTextFile("C:\R-Tel  Tester\Input-ZTE-64p-Attn-SNR-Check.txt", ForReading, False)
	'Set file = fso.OpenTextFile("D:\Raj-Tel Tester\Input-ZTE-64p-Attn-SNR-Check.txt", ForReading, False)

  
 Do While file.AtEndOfStream <> True
           str = file.Readline
           crt.Screen.Send "telnet" & " " & str & vbcr
 
          row=nxtrow
              
          	If Not crt.Screen.waitforstring ("ogin:", 1) Then	 '------------------ Initial Enter in ZTE---------------------
		crt.Screen.Send vbCR
	Else
		crt.Screen.Send vbCR
	End If  

     If crt.Screen.waitforstring ("ogin:", 9 ) Then  '------------------If DSLAM is Down / No Telnet then wait for 9 Seconds ------------------------
           	crt.Screen.Send "admin" & vbcr
           	crt.Screen.waitforstring ("ssword:")
           	crt.Screen.Send "admin" & vbcr
           	crt.Screen.waitforstring (">") 

	Dim screenrow, readline, chakra1, chakra2, ajay1, ajay2, looper1, looper2, looper3, nakal1, nakal2, asli1, asli2, portkram, sodh1
		zkram = nextzkram
		ws.Cells(row, 1).Value =  zkram
		ws.Cells(row, 2).Value =  Date
		ws.Cells(row, 3).Value =  Time
      		ws.Cells(row, 4).Value = str

		crt.Screen.Send "enable" & vbcr
		crt.Screen.waitforstring ("ssword:") 
		crt.Screen.Send "admin" & vbcr
		crt.Screen.waitforstring ("#")

		screenrow = crt.screen.Currentrow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 35 )
		ws.Cells(row, 5).Value = readline

  		crt.Screen.Send "show card" & vbcr
  		crt.Screen.waitforstring ("#")


		ajay2 = 3
		looper1 = 0
		Do
			screenrow = crt.screen.Currentrow - ajay2
			readline = crt.Screen.Get(screenrow, 3, screenrow, 6 )
			ajay1 = readline

			If ajay1 = "----" Then Exit Do

			readline = crt.Screen.Get(screenrow, 62, screenrow, 66 )
			nakal1 = readline
			If nakal1 = "Inser" then
				readline = crt.Screen.Get(screenrow, 22, screenrow, 23 )
				sodh1 = readline
				If sodh1 > 15 Then
					readline = crt.Screen.Get(screenrow, 3, screenrow, 3 )
					zteslot(looper1) = readline
					ztecard(looper1) = sodh1
					looper1 = looper1 + 1
				end If
				
			End If

			ajay2=ajay2+1

		Loop

		asli1 = looper1 - 1

		Dim barfi, ninjasnr, ninjaattn, ninjasnr2, ninjaattn2

		For looper2 = asli1 to 0 step -1

			For looper3 = 1 to ztecard(looper2) step + 1

				crt.Screen.Send "show adsl physical " & zteslot(looper2) & "/" & looper3 & vbcr

        				If crt.Screen.waitforstring ("#", 3) Then
					ws.Cells(row, 6).Value = "Port " & zteslot(looper2) & "/" & looper3
					ws.Cells(row, 7).Value = "Down"
					row = row + 1
				Else
					ws.Cells(row, 6).Value = "Port " & zteslot(looper2) & "/" & looper3
					ws.Cells(row, 7).Value = "Up"

					screenrow = crt.screen.Currentrow - 48
					'readline = crt.Screen.Get(screenrow, 9, screenrow, 11 )	
					'ninjasnr = readline

					'If ninjasnr = "Snr" Then
					readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'------------- ATU-C SNR -----------
					ws.Cells(row, 12).Value = readline
					'Else
						'screenrow = crt.screen.Currentrow - 47
						'readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'------------- ATU-C SNR -----------
						'ws.Cells(row, 12).Value = readline
					'End If

					screenrow = crt.screen.Currentrow - 47
					readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'------------------ATU-C Attenua.-----------
					ws.Cells(row, 14).Value = readline
					
					screenrow = crt.screen.Currentrow - 45
					readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'---------------- ATU-C O/p power-----------
					ws.Cells(row, 18).Value = readline
					
					screenrow = crt.screen.Currentrow - 44
					readline = crt.Screen.Get(screenrow, 27, screenrow, 33 )		'----------------- ATU-C Attain rate ----------
					ws.Cells(row, 16).Value = readline

					screenrow = crt.screen.Currentrow - 18
					readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'------------- ATU-R SNR -----------
					ws.Cells(row, 13).Value = readline

					screenrow = crt.screen.Currentrow - 17
					readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'------------------ATU-R Attenua.-----------
					ws.Cells(row, 15).Value = readline
					
					screenrow = crt.screen.Currentrow - 15
					readline = crt.Screen.Get(screenrow, 27, screenrow, 30 )		'---------------- ATU-R O/p power-----------
					ws.Cells(row, 19).Value = readline
					
					screenrow = crt.screen.Currentrow - 14
					readline = crt.Screen.Get(screenrow, 27, screenrow, 33 )		'----------------- ATU-R Attain rate ----------
					ws.Cells(row, 17).Value = readline
					
					barfi = row

					ws.Cells(row, 8).Value = "=L" & barfi & "/10"
					ws.Cells(row, 9).Value = "=M" & barfi & "/10"
					ws.Cells(row, 10).Value = "=N" & barfi & "/10"
					ws.Cells(row, 11).Value = "=O" & barfi & "/10"

					row = row + 1
					crt.Screen.Send  vbcr
					crt.Screen.waitforstring ("quit)")
					'If not crt.Screen.waitforstring ("#", ) Then
					crt.Screen.Send  vbcr
					'End If
					crt.Screen.waitforstring ("#")
				End If
			Next
		Next	
		
		For looper2 = asli1 to 0 step -1
			zteslot(looper2) = 0
			ztecard(looper2) = 0
		Next

		crt.Screen.Send "quit" & vbcr
		crt.Screen.waitforstring (":[N]")
		crt.Screen.Send "y" & vbcr

		nxtrow = row
		zkram = zkram + 1
		nextzkram = zkram


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


   'crt.Screen.WaitForString "mnx-01>"		'------------------Surat BNG

loop

    '----------------------------------------------------------------  Loop - Now Test second IP in Input List -----------------------

    wb.SaveAs("H:\R-Tel  Tester\Raj-ZTE-64p-ATTN-SNR-result.xls")
    'wb.SaveAs("C:\R-Tel  Tester\Raj-ZTE-64p-ATTN-SNR-result.xls")

    'wb2.SaveAs("H:\R-Tel  Tester\Raj-TEL-UT-result.xls")

    'wb.SaveAs("D:\Raj-Tel Tester\Raj-ZTE-64p-ATTN-SNR-Result.xls")
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

'-----------------------------------------------((((((((    ZTE - 64p DSLAM SNR - Attenuatio Check Script   ))))))------------------------------------------------------
