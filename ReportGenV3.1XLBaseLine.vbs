Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim oFSO, oFolder, oFile, oTextStream, strText, fNameReport, oXMLHttp, ohtmlFile

'Dim violationsTotal
'violationsTotal = 0
'Dim incompleteTotal
'incompleteTotal = 0

Set oIE = CreateObject("InternetExplorer.Application")
Set oIE2 = CreateObject("InternetExplorer.Application")

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder("AxeComplianceReports")
Set filetoWrite = oFSO.GetFolder(oFolder.ParentFolder)
Set xdoc = CreateObject("excel.application")

xdoc.Application.Visible = true 
xdoc.Workbooks.Add
xdoc.Cells(1,1).value = "Page"
xdoc.Cells(1,2).value = "Issues"
xdoc.Cells(1,3).value = "Owner"

Dim keyValuePairs(79)
keyValuePairs(0)="AboutPage=Retail"
keyValuePairs(1)="HomePage=Retail"
keyValuePairs(2)="AccessibilityPage=Core"
keyValuePairs(3)="IDPage=ID"
keyValuePairs(4)="LoggingInWithUnidaysPage=ID"
keyValuePairs(5)="LoginPage=ID"
keyValuePairs(6)="NotificationsPage=Core"
keyValuePairs(7)="SettingsPage=ID"
keyValuePairs(8)="StudentStatusPage=ID"
keyValuePairs(9)="AllPage=Retail"
keyValuePairs(10)="BeautyAllBeautyPage=Retail"
keyValuePairs(11)="BeautyFragrancePage=Retail"
keyValuePairs(12)="BeautyHairPage=Retail"
keyValuePairs(13)="BeautyMakeUpPage=Retail"
keyValuePairs(14)="BeautyPage=Retail"
keyValuePairs(15)="BeautyPremiumSkincarePage=Retail"
keyValuePairs(16)="BecomeAnAmbassadorPage=Retail"
keyValuePairs(17)="CareersPage=Core"
keyValuePairs(18)="ContactPage=Creative-Commercial"
keyValuePairs(19)="CookiePolicyPage=Core"
keyValuePairs(20)="CorporateContactPage=Creative-Commercial"
keyValuePairs(21)="CorporateGenZInsightsPage=Creative-Commercial"
keyValuePairs(22)="CorporateMarketplaceBeautyPage=Creative-Commercial"
keyValuePairs(23)="CorporateMarketplaceFashionPage=Creative-Commercial"
keyValuePairs(24)="CorporateMarketplaceFoodTravelLeisurePage=Creative-Commercial"
keyValuePairs(25)="CorporateMarketplacePage=Creative-Commercial"
keyValuePairs(26)="CorporateMarketplaceTechPage=Creative-Commercial"
keyValuePairs(27)="CorporateMarketplaceWellnessEarningLearningPage=Creative-Commercial"
keyValuePairs(28)="CorporatePage=Creative-Commercial"
keyValuePairs(29)="CorporateStudentVerificationPage=Creative-Commercial"
keyValuePairs(30)="FashionAccessoriesPage=Retail"
keyValuePairs(31)="FashionAllFashionPage=Retail"
keyValuePairs(32)="FashionClothingPage=Retail"
keyValuePairs(33)="FashionJewelleryAndWatchesPage=Retail"
keyValuePairs(34)="FashionLingerieAndUnderwearPage=Retail"
keyValuePairs(35)="FashionMenswearPage=Retail"
keyValuePairs(36)="FashionPage=Retail"
keyValuePairs(37)="FashionShoesPage=Retail"
keyValuePairs(38)="FoodAndDrinkAllFoodPage=Retail"
keyValuePairs(39)="FoodAndDrinkDineOutPage=Retail"
keyValuePairs(40)="FoodAndDrinkHomeCookingPage=Retail"
keyValuePairs(41)="FoodAndDrinkPage=Retail"
keyValuePairs(42)="FoodAndDrinkTakeawayAndDeliveryPage=Retail"
keyValuePairs(43)="HealthAndFitnessAllHealthAndFitnessPage=Retail"
keyValuePairs(44)="HealthAndFitnessClothingPage=Retail"
keyValuePairs(45)="HealthAndFitnessEyeCarePage=Retail"
keyValuePairs(46)="HealthAndFitnessGymPage=Retail"
keyValuePairs(47)="HealthAndFitnessHealthSubscriptionsPage=Retail"
keyValuePairs(48)="HealthAndFitnessPage=Retail"
keyValuePairs(49)="HealthAndFitnessSupplementsPage=Retail"
keyValuePairs(50)="HealthAndFitnessTrackersPage=Retail"
keyValuePairs(51)="CookieSettingsPage=Core"
keyValuePairs(52)="JoinPage=ID"
keyValuePairs(53)="LearningAndEarning_EarningPage=Growth"
keyValuePairs(54)="LearningAndEarning_LanguagesPage=Growth"
keyValuePairs(55)="LearningAndEarning_PersonalGrowthPage=Growth"
keyValuePairs(56)="LearningAndEarning_StudyPage=Growth"
keyValuePairs(57)="LearningAndEarning_WellbeingPage=Growth"
keyValuePairs(58)="LifestyleAllLifestylePage=Retail"
keyValuePairs(59)="LifestyleBooksAndStationeryPage=Retail"
keyValuePairs(60)="LifestyleEntertainmentPage=Retail"
keyValuePairs(61)="LifestyleHolidaysAndHotelsPage=Retail"
keyValuePairs(62)="LifestyleHomeAndBankingPage=Retail"
keyValuePairs(63)="LifestylePage=Retail"
keyValuePairs(64)="LifestyleSubscriptionsPage=Retail"
keyValuePairs(65)="LifestyleTravelAndTransportPage=Retail"
keyValuePairs(66)="LimitedTimeOnlyPage=Retail"
keyValuePairs(67)="ResetPasswordPage=ID"
keyValuePairs(68)="ASOSPartnerPage=Retail"
keyValuePairs(69)="PressPage=Creative-Commercial"
keyValuePairs(70)="PrivacyPolicyPage=Core"
keyValuePairs(71)="SupportPage=Retail-ID"
keyValuePairs(72)="TechnologyAccessoriesPage=Retail"
keyValuePairs(73)="TechnologyAllTechnologyPage=Retail"
keyValuePairs(74)="TechnologyElectricalsPage=Retail"
keyValuePairs(75)="TechnologyGamingPage=Retail"
keyValuePairs(76)="TechnologyLaptopsAndTabletsPage=Retail"
keyValuePairs(77)="TechnologyMobilePage=Retail"
keyValuePairs(78)="TechnologyPage=Retail"
keyValuePairs(79)="TermsOfServicePage=Core"


rowCount = 2
For Each oFolder In oFolder.SubFolders
	toNodeDet = ""
	If oFolder.name <> "V5Accessibility_HomePageAndCookies" Then	
		
		Set colFiles = oFolder.Files	
				
		For Each objFile in colFiles
				If objFile.name = "UNiDAYS_HomePage.html" Then
				Else
					With oIE
						.Visible = false
						.Navigate oFSO.GetAbsolutePathName (objFile.path)
					
						Do While .Busy Or .ReadyState <> 4
							WScript.Sleep 10
						Loop
						
						Set oNode = oIE.Document.querySelectorAll("div.emOne")
						Set oNodeDetails = oIE.Document.getElementsByClassName("emTwo")
							For Each oNodeDet In oNodeDetails							
								divElementchunk = Split(oNodeDet.innerText,"<br>",-1,1)
								For Each workChunk In divElementchunk									
									If InStr(workChunk,"Impact: seri") > 0 Or InStr(workChunk,"Impact: crit") Then
										secondLayer = Split(workChunk,vbCrLf)
										For Each slInfo In secondLayer
											If InStr(slInfo,"Tags:") > 0 Then
												tagsSplit = Split(slInfo,",")
												levelInfo = ""
												For iterCol = LBound(tagsSplit)+1 To UBound(tagsSplit)'													
													If iterCol = LBound(tagsSplit)+1 Then
														levelInfo = Trim(StrOperator(tagsSplit(iterCol)))
													Else
														toXL = levelInfo + StrOperator(tagsSplit(iterCol))
														If InStr(toXL,"-") <> 0 Then
															valToExcel = Mid(oFSO.GetBaseName(objFile.Name),9)														
															xdoc.Cells(rowCount,1).value = Mid(oFSO.GetBaseName(objFile.Name),9)
															xdoc.Cells(rowCount,2).value = toXL
															xdoc.Cells(rowCount,3).value = CheckKeyValuePair(keyValuePairs, valToExcel)
															rowCount = Int(rowCount) + 1
														Else
														End If	
													End If																										
												Next												
											Else
											End If	
										Next	
									Else										
									End If									
								Next
							Next
					End With				
				End If
				Set oNodeDetails =	Nothing
		Next
			
	Else
		Set colFiles = oFolder.Files	
				
		For Each objFile in colFiles
			With oIE2
				.Visible = false
				.Navigate oFSO.GetAbsolutePathName (objFile.path)
			
				Do While .Busy Or .ReadyState <> 4
					WScript.Sleep 10
				Loop
				
					Set oNode = oIE2.Document.querySelectorAll("div.emOne")						
					Set oNodeDetails = oIE2.Document.getElementsByClassName("emTwo")
							For Each oNodeDet In oNodeDetails							
								divElementchunk = Split(oNodeDet.innerText,"<br>",-1,1)
								For Each workChunk In divElementchunk									
									If InStr(workChunk,"Impact: seri") > 0 Or InStr(workChunk,"Impact: crit") Then
										secondLayer = Split(workChunk,vbCrLf)
										For Each slInfo In secondLayer
											If InStr(slInfo,"Tags:") > 0 Then
												tagsSplit = Split(slInfo,",")
												levelInfo = ""
												For iterCol = LBound(tagsSplit)+1 To UBound(tagsSplit)'													
													If iterCol = LBound(tagsSplit)+1 Then
														levelInfo = Trim(StrOperator(tagsSplit(iterCol)))
													Else
														toXL = levelInfo + StrOperator(tagsSplit(iterCol))
														If InStr(toXL,"-") <> 0 Then
															valToExcel = Mid(oFSO.GetBaseName(objFile.Name),9)														
															xdoc.Cells(rowCount,1).value = Mid(oFSO.GetBaseName(objFile.Name),9)
															xdoc.Cells(rowCount,2).value = toXL															
															xdoc.Cells(rowCount,3).value = CheckKeyValuePair(keyValuePairs, valToExcel)
															rowCount = Int(rowCount) + 1
														Else
														End If	
													End If																										
												Next												
											Else
											End If	
										Next	
									Else										
									End If									
								Next
							Next
			End With				
		Next
	End If	
Next

Function GetDateTimeStamp
  Dim strNow
  strNow = Now()
  GetDateTimeStamp = Year(strNow) & Pad2(Month(strNow)) _
        & Pad2(Day(StrNow)) & Pad2(Hour(strNow)) _
        & Pad2(Minute(strNow)) & Pad2(Second(strNow))
End Function

Function Pad2(strIn)
  Do While Len(strIn) < 2
    strIn = "0" & strIn
  Loop
  Pad2 = strIn
End Function

Function StrOperator(inp)
	Select Case Trim(inp)
		Case "wcag2a"
			OutPut = Replace(Trim(inp),"wcag2a","Level A")
			StrOperator = OutPut
		Case "wcag2aa"
			OutPut = Replace(Trim(inp),"wcag2aa","Level AA")
			StrOperator = OutPut		
		Case Else			
			Set re = New RegExp
			With re
				.Pattern = "^wcag\d"
				.IgnoreCase = False				
			End With	
			If re.Test(Trim(inp)) Then
				strRep2 = Replace(Trim(inp),"wcag","wcag ")
				strOp = Mid(strRep2,InStr(strRep2," "))	
				If strOp <> "" Then 
				    For i = 2 To Len(strOp)
				    	If i=Len(strOp) Then
				    		OutPut = OutPut &  Mid(strOp,i,1) 
				    	Else	 
				    		OutPut = OutPut &  Mid(strOp,i,1) & "." 
				    	End If
				    Next     
				End If
				StrOperator = " - "&OutPut
			Else
			End If
			Set re = Nothing
	End Select	
End Function

Set objWorksheet = xdoc.Worksheets(1)
objWorksheet.name = "a11yReport"

XLPivotTable(objWorksheet.name)

Function XLPivotTable(workSheetName)

	Set objData = xdoc.Worksheets(workSheetName)
	Set objActiveSheet = xdoc.ActiveSheet
	Const xlR1C1 = -4150
	'Const xlA1 = 1
	objActiveSheet.Columns.Autofit
	SrcData = workSheetName&"!" & objData.UsedRange.Address(xlR1C1)
	
	Const xlDatabase = 1
	Set pvtTable = xdoc.ActiveWorkbook.PivotCaches.Create(xlDatabase,SrcData).CreatePivotTable(workSheetName&"!R2C5","PivotTable1")
		
	Const xlRowField = 1
	
	Const xlColumnField = 2	
	
	Set pagePivot = pvtTable.pivotFields("Page")
	pagePivot.orientation = xlRowField
		
	Set issuesPivot = pvtTable.pivotFields("Issues")
	issuesPivot.orientation = xlRowField
	
	Set ownerPivot = pvtTable.pivotFields("Owner")
	ownerPivot.orientation = xlColumnField
	
	Const xlCount = -4112
	Const xlSum = -4157
		
	Const xlDataField = 4
	Set valueFieldPivot = pvtTable.PivotFields("Issues")    
    valueFieldPivot.Orientation = xlDataField
    valueFieldPivot.Function = xlCount
	
	FormatPivotTableStyle(pvtTable)
	
			
End Function

Sub FormatPivotTableStyle(pivotTable)
    If IsNull(pivotTable) Then
        Err.Raise vbObjectError + 1, , "Pivot table object is required."
    End If    

    pivotTable.TableStyle2 = "PivotStyleMedium9"
    
    pivotTable.ShowTableStyleRowStripes = True
    pivotTable.ShowTableStyleColumnStripes = True
    pivotTable.ShowTableStyleLastColumn = True
    pivotTable.ShowTableStyleRowHeaders = False
    pivotTable.ShowTableStyleColumnHeaders = True
    pivotTable.ShowTableStyleRowStripes = True
    pivotTable.ShowTableStyleColumnStripes = True
End Sub

Function CheckKeyValuePair(keyValuePairs, valueToCheck)    
    For Each pair In keyValuePairs        
        pairArray = Split(pair, "=")
        key = pairArray(0)
        value = pairArray(1)        
        
        If key = valueToCheck Then
            CheckKeyValuePair = value
            Exit Function
        End If
    Next 
    
    CheckKeyValuePair = ""
End Function

xdoc.Application.ActiveWorkbook.SaveAs oFSO.GetAbsolutePathName(filetoWrite)&"\"&"a11yReport_"&GetDateTimeStamp&".xlsx"
xdoc.Application.Visible = false
xdoc.Application.Quit
Set objData = Nothing
Set objActiveSheet = Nothing
Set pvtTable = Nothing
Set objWorksheet = Nothing
Set xdoc = Nothing
Set oNode = Nothing
Set oNodeDetails = Nothing
Set oFSO = Nothing
Set oIE = Nothing
Set oIE2 = Nothing