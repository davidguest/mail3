This version of the mail facility is designed for use with Office 365 and uses Exchange Web Services to communicate with the server.

The server pages are written in PHP and handle all of the communication with Office 365. The client page is a javascript app designed for use as an App Extension Kit screen in the CampusM mobile app. 

##Installation instructions
1. put the contents of the server folder on to your web server at an appropriate address
2. set up a web service endpoint in the CampusM management portal
3. add a new AEK page with the contents of aek.html
4. change the webservice_url parameter on line 73 of the AEK page to point to the mail.php page on your web server
5. make sure the web service endpoint you created in step 2 is referenced in line 77 of the AEK page
6. change line 81 of the AEK page with the appropriate suffix for your users
