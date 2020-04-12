SystemUtil.Run "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe","","C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\",""
SystemUtil.Run "C:\Users\thispc\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Flight GUI.lnk","","",""
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set "john" @@ hightlight id_;_1912817680_;_script infofile_;_ZIP::ssf2.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "HP" @@ hightlight id_;_2105029152_;_script infofile_;_ZIP::ssf3.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_1951272176_;_script infofile_;_ZIP::ssf4.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").Output CheckPoint("John Smith") @@ hightlight id_;_2084043760_;_script infofile_;_ZIP::ssf5.xml_;_
