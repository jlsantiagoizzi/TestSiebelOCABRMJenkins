
Call FnLogin()

Function FnLogin()
n=0

	Do 
		systemutil.CloseProcessByName "iexplore.exe"
		systemutil.CloseProcessByName "SiebelAx_Test_Automation_21233.exe"
		systemutil.Run "iexplore.exe","http://172.21.40.87/ecommunications_esn/start.swe?SWECmd=AutoOn"
		
		Browser("Siebel Communications").Page("Siebel Communications").WebEdit("_SweUserName").Set "CVSUCROBOT"
		Browser("Siebel Communications").Page("Siebel Communications").WebEdit("_SwePassword").SetSecure "Shenlong12"
		Browser("Siebel Communications").Page("Siebel Communications").Image("LoginButton").Click
		wait 15
		
	    n=n+1
	    If n>5 Then 
	        print "Siebel no funciona"
	    	ExitActionIteration
	    End If
	    
	Loop while not SiebApplication("Siebel Communications").SiebPageTabs("PageTabs").Exist	
	
End Function


