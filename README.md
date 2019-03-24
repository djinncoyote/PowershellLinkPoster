# PowershellLinkPoster
Powershell Link Poster - Pinboard to Dreamwidth

Note: This script uses the Dreamwidth "Post by Email" functionality.  It runs under PowerShell, which comes as default in Windows 10.  
The script may not work for PowerShell running on MacOS or Linux, due to its use of a system environment variable.

Run with a command-line like:

`postlinks.ps1 -pinboardUser andrewducker -emailFrom andrew@ducker.org.uk -emailTo andrewducker+1234@post.dreamwidth.org`

You will need to set up post by email at https://www.dreamwidth.org/manage/emailpost and then update the command line to match your settings.

You can run it with a TestMode flag, which will show the results rather than posting them.

i.e. `postlinks.ps1 -pinboardUser andrewducker -emailFrom andrew@ducker.org.uk -emailTo andrewducker+1234@post.dreamwidth.org -TestMode`

If you're inside a corporate firewall then you will probably need to set the ProxyCredentials parameter.  And unless you have external SMTP access you'll need an internal mail server (use the "smtpServer" parameter).

The first time the script is run, it will fetch all links that are present on your RSS feed - Pinboard keeps up to 70 of your last pinned links in the feed.
The script will also set an environment variable to store the date and time that the script was last run.

For subsequent runs, the script will fetch all links posted since the last time the script was run, and will update the environment variable each time.

You can use the "dateFrom" and/or "dateTo" parameters to override which links are fetched (when "dateTo" is used, the environment variable will be set to that value):

`postlinks.ps1 -pinboardUser andrewducker -emailFrom andrew@ducker.org.uk -emailTo andrewducker+1234@post.dreamwidth.org` -dateFrom 2019-03-20T12:00:00 -dateTo 2019-03-22T12:00:00

All suggestions/merge requests gratefully received.
