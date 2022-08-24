 
             There are 2 parts to this alarm program
            
             Daily report on the state of the system
             This reflects the state of the system (all settop units) at the time of the email
             The main thing is that the email comes out once a day regardless of alarm / no alarm state
             This confirms that the server is operating and that the alarm email check is working
            
             Data flow from the Settop units
             This checks the time of the last data file received for each connected Settop
             If it is older than the defined time, then the first of 3 alarm emails is sent out
             This is repeated for 2 more emails
             The alarm emails then stop and you only detect that the system has a faulty settop by looking at the daily status email.
             This stiops hundreds of emails being sent
             
             Can monitor up to 5 settops at a time - increase by adding more folders
             Problem is that all folders are subject to the same data interval 
             only works if they are running at approximately the same frequency
             The gka folder is the folder that contains the date sub folders
            
             if testEmail "Yes", then a test email is sent to confirm that the email functions.



             Julian Gray
             gna.geomatics@gmail.com
             +49 176 7299 7904
            
