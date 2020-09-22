Before using this program please back up your hosts file. Please check the hosts.svr file(in data directory) for names of servers you might not want to block.

Please do not add too many server names in the blocked list, because an inflated hosts file may actually slow down your internet connection. (around 600 server names are already there in the blocked list, but try to bring it below 500 by customizing the list). Also, please don't add too many server names and their ip addresses in the DNS Cache because that would again get added up in the hosts file. Only add ips & site names of sites you visit frequently. The format should be -
ip address <space> sitename eg:- 216.5.163.42 www.planetsourcecode.com

If you want to enble selective cookie deletion, then right click and tray icon and go to settings. Click on the cookies tab and add the name of the server from which no cookie is to be deleted. For eg: You don't want to delete any cookies from Planetsourcecode.com, so what you do is - add planetsourcecode to the list, and uncheck "delete all cookies" option.

Reducing the timer interval can cause unpredictable behaviour, atleast in terms of blocking popups. The popup blocked count show in the systray will almost always be wrong and valid popups will also get closed. Besides, it will also be a wastage of system resouces. Please experiment with your system settings to get optimum output. It would be safer to disable popup blocking while browsing sites where you would rather want to see popups (eg: image galleries etc.)

All files in Data & Image are required for proper working of the program and it should be present in the application directory.

Please refer to Readme.WRI for details.

Please report bugs, if you find any.Thanks for downloading - Enjoy.


NOTE:

If you don't want to use selective popup blocking, you may well replace the code in the timer proc with the following :-

If mnublockpopups.Checked = True Then
    'get foreground window handle
    hnd = GetForegroundWindow()
     
    'see whether it an IE window ( IEFRAME class for IE 5.5 and CabinetWClass for IE 4-5\)
    Call GetClassName(hnd, class, 15)
    If Left(class, 7) = "IEFrame" Or Left(class, 13) = "CabinetWClass" Then
       
        if the window is visible then find its first child window
        If IsWindowVisible(hnd) <> False Then
            hcld = FindWindowEx(hnd, GW_CHILD And GW_HWNDFIRST, vbNullString, vbNullString)
            
            'if the 1st child window WorkerA (top of zorder) is invisible, that means this window has not   
            'toolbars etc and it is a popup, so close it
            Call GetClassName(hcld, class, 8)
                       
            If Left(class, 7) = "WorkerA" And IsWindowVisible(hcld) = False Then
                Call PostMessage(hnd, WM_CLOSE, 0, 0)
                stats(2) = stats(2) + 1
            End If

        End If

    End If

End If
(only the popup blocking code section)
This will result in efficient blocking of popups (the timer interval may be safely set to 500)

/data/cookies.svr - names of sites from which cookies are not to be deleted
/data/hosts.svr - names of blocked sites
/data/hosts.bak - backup of original hosts file
/data/popups.svr - captions of popups not to be closed
/data/dns.svr - ip ads and server names in DNS Cache
/images/*.gif - web server image files
/images/*.ico - iconsused in this project
