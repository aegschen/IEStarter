	## IEStarter (Opens some IE windows on three monitors.)
	## Version: 1.1 (30.01.2017)

Function Wait_IE ($ieX) { while($ieX.Busy -eq $true){Start-Sleep -Milliseconds 5000} }

	# Monitor 1 IE-Frame 1
	if (!$site_addrA)
	{
	$site_addrA = "http://www.ge1-nb.ch/einsatzpl/show_list.php?filter1=Pikettdienst&filter2=&button2=Filter+setzen&filter4=&filter5=&feld1=2016-09-20"
	}
	
	# Monitor 1 IE-Frame 2
	if (!$site_addrB)
	{
	$site_addrB = "http://www.google.ch"
	}
	
	# Monitor 2 IE-Frame 1
	if (!$site_addrC)
	{
	$site_addrC = "http://www.ge1-nb.ch/einsatzpl/show_list.php?filter1=Reinigung&filter2=&button2=Filter+setzen&filter4=&filter5=&feld1=2016-09-20"
	}
	
	# Monitor 2 IE-Frame 2
	if (!$site_addrD)
	{
	$site_addrD = "http://www.microsoft.com"
	}
    
	# Monitor 3 IE-Frame 1
	if (!$site_addrE)
	{
	$site_addrE = "http://www.ge1-nb.ch/einsatzpl/show_list.php?filter1=Gr%FCnpflege&filter2=&button2=Filter+setzen&filter4=&filter5=&feld1=2016-09-20"
	}
	
	# Monitor 3 IE-Frame 2
	if (!$site_addrF)
	{
	$site_addrF = "http://www.20min.ch"
	}
	
 

	$ieA = New-Object -com InternetExplorer.Application
    ## Sets the size of the window
	$ieA.Top = 0
	$ieA.Left = 0
	$ieA.Width = 1080
	$ieA.Height = 960

	## Turns off the unnecessary menus and tools and don't sets the window as resizable
	$ieA.AddressBar = $false
	$ieA.MenuBar = $false
	$ieA.ToolBar = $false
	$ieA.Resizable = $false
	$ieA.StatusBar = $false

	$ieB = New-Object -com InternetExplorer.Application
    ## Sets the size of the window
	$ieB.Top = 960
	$ieB.Left = 0
	$ieB.Width = 1080
	$ieB.Height = 960

	## Turns off the unnecessary menus and tools and don't sets the window as resizable
	$ieB.AddressBar = $false
	$ieB.MenuBar = $false
	$ieB.ToolBar = $false
	$ieB.Resizable = $false
	$ieB.StatusBar = $false

	$ieC = New-Object -com InternetExplorer.Application
    ## Sets the size of the window
	$ieC.Top = 0
	$ieC.Left = 1080
	$ieC.Width = 1080
	$ieC.Height = 960

	## Turns off the unnecessary menus and tools and don't sets the window as resizable
	$ieC.AddressBar = $false
	$ieC.MenuBar = $false
	$ieC.ToolBar = $false
	$ieC.Resizable = $false
	$ieC.StatusBar = $false

	$ieD = New-Object -com InternetExplorer.Application
    ## Sets the size of the window
	$ieD.Top = 960
	$ieD.Left = 1080
	$ieD.Width = 1080
	$ieD.Height = 960

	## Turns off the unnecessary menus and tools and don't sets the window as resizable
	$ieD.AddressBar = $false
	$ieD.MenuBar = $false
	$ieD.ToolBar = $false
	$ieD.Resizable = $false
	$ieD.StatusBar = $false
	
	$ieE = New-Object -com InternetExplorer.Application
    ## Sets the size of the window
	$ieE.Top = 0
	$ieE.Left = 2160
	$ieE.Width = 1080
	$ieE.Height = 960

	## Turns off the unnecessary menus and tools and don't sets the window as resizable
	$ieE.AddressBar = $false
	$ieE.MenuBar = $false
	$ieE.ToolBar = $false
	$ieE.Resizable = $false
	$ieE.StatusBar = $false
	
	$ieF = New-Object -com InternetExplorer.Application
    ## Sets the size of the window
	$ieF.Top = 960
	$ieF.Left = 2160
	$ieF.Width = 1080
	$ieF.Height = 960

	## Turns off the unnecessary menus and tools and don't sets the window as resizable
	$ieF.AddressBar = $false
	$ieF.MenuBar = $false
	$ieF.ToolBar = $false
	$ieF.Resizable = $false
	$ieF.StatusBar = $false
	

$date = Get-Date

While ($date -ge "1/1/16" -and $date -le "1/1/2030")
    
    {	$ieA.Visible = $true
        $ieA.Navigate("$site_addrA")

        Wait_IE $ieA
		
		$ieB.Visible = $true
        $ieB.Navigate("$site_addrB")
                
        Wait_IE $ieB

		$ieC.Visible = $true
        $ieC.Navigate("$site_addrC")

        Wait_IE $ieC 
		
		$ieD.Visible = $true
		$ieD.Navigate("$site_addrD")

        Wait_IE $ieD

		$ieE.Visible = $true
		$ieE.Navigate("$site_addrE")

        Wait_IE $ieE

		$ieF.Visible = $true
		$ieF.Navigate("$site_addrF")

        Wait_IE $ieF 

        $ieA.Refresh()
        $ieB.Refresh()
        $ieC.Refresh()
		$ieD.Refresh()
		$ieE.Refresh()
		$ieF.Refresh()

        # Set sleep time for refresh   
        Start-Sleep -Seconds 10
                    
    }

