

Dim StrConfigFlagStat,StrConfigTcId,StrRowCntDriver,StrDriverTcId,StrColCntDriver,StrFuncName

 @@ hightlight id_;_1909830256_;_script infofile_;_ZIP::ssf35.xml_;_
 @@ hightlight id_;_1909846528_;_script infofile_;_ZIP::ssf33.xml_;_
 @@ hightlight id_;_1909846672_;_script infofile_;_ZIP::ssf34.xml_;_
Datatable.AddSheet("Config") 
Datatable.AddSheet("Driver") ' Adding the Driver Sheet in the Datatable
Datatable.AddSheet("TestData") 'Adding the TestData Sheet in the Datatable

Datatable.ImportSheet "C:\Users\thispc\Documents\Unified Functional Testing\man\Config\Config.xlsx", "Config", "Config" 'Importing he Config Sheet in the Datatable
Datatable.ImportSheet "C:\Users\thispc\Documents\Unified Functional Testing\man\Config\Driver.xlsx", "Driver", "Driver" ' Importing the Driver Sheet in the Datatable
'Datatable.ImportSheet "C:\Users\thispc\Documents\Unified Functional Testing\man\TestData\TestData.xlsx", "TC1", "TestData" 'Importing the TestData Sheet in the Datatable

introw = datatable.GetSheet("Config").GetRowCount ''Getting the rowcount of the Config Sheet

For i =1 to introw step 1'Loop for controlling the Config  

	Datatable.GetSheet("Config").SetCurrentRow(i)	'Setting the Current row of the Config sheet  from the loop initializing from 1 - Remember - the cursor will be set for the config and Driver File.

	StrConfigFlagStat = datatable.GetSheet("Config").GetParameter("Execute") ' Getting the "Execute" column value of the Config file i.e the test cases whose flag values set to "Yes"

	If StrConfigFlagStat = "Yes" Then 'Condition to execute  only those test cases which is marked as "Yes"

		StrConfigTcId = datatable.GetSheet("Config").GetParameter("TestCaseId") 'if the above condition is true then getting the name of the Test cases which is marked as "Yes" in the config file.

		StrRowCntDriver=datatable.GetSheet("Driver").GetRowCount 'Getting the total row count of the Driver sheet
						
						For j =1 to StrRowCntDriver 'Loop for controlling the Driver File
						
											datatable.GetSheet("Driver").SetCurrentRow(j)

											StrDriverTcId  = datatable.GetSheet("Driver").GetParameter("TestCaseId").Value  'Getting the "TestCaseId" column of the Driver file.
												
												'Msgbox(a)
											
												If StrDriverTcId = StrConfigTcId Then 'Condition to check whehter the Driver file TCID is matching with the Config File TCID.
					
													Environment.Value("StrTestCase") = StrConfigTcId 'If above condition is true then Storing the Config file TCID which is marked for execution into the environment variable- So that it can be used  for retrieving the data from Test data  sheet  which is marked for execution in the config file.
										
													StrColCntDriver= Datatable.GetSheet("Driver").GetParameterCount 'Getting the total column count of the Driver sheet
													
															For k=1 to StrColCntDriver step 1 'Loop for executing the number of function for a particular test case in the Driver File. "StrColCntDriver-1" because the first colum is the Test Case id in the driver file
																
																	StrFuncName = datatable.GetSheet("Driver").GetParameter(k).Value 'Getting the function name from the driver file.

																	If  StrFuncName = "Yes" Then

																			a = Datatable.GetSheet("Driver").GetParameter(k).Name
																			
																	
																			ExecuteGlobal a

																	End If
					
																	' Executing the Function. - Note - If StrFuncName is even blank i.e "" then also it will execute and will not find anything in the function library.Advantage it will  not throw any error even if it's dosent found anything in the function library. 
													
																	'If StrFuncName="" Then ' Condition for checking whether the  String is blank
							
																	'		Exit For 'If above condition is True then exit  for reading the function from the driver file.
							
																	'End If
													
															Next 'Loop  -  for executing the number of function for a particular test case in the Driver File
											
												End If  'End - Condition to check whehter the Driver file TCID is matching with the Config File TCID.
						
									Datatable.SetNextRow ' Incrementing the cursor to the next row of the Driver datatable.
						
						Next 'Loop  - for controlling the Driver File

			Datatable.SetNextRow ' Incrementing the cursor to the next row of the Config datatable.
	
	End If 'End - Condition to execute  only those test cases which is marked as "Yes"


Next 'Loop  - for controlling the Config File

exittest @@ hightlight id_;_1963059168_;_script infofile_;_ZIP::ssf10.xml_;_
