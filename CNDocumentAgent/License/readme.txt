It will search for cndalicense.lic in app folder.

if not found, it will assume trial and create a trial license (cnda.dat) in the input\processed folder. 
This stores the date which the app will then read each time to see if it exceeded the 30 day trial license. 

the cndalicense.lic file, position 5 determines what kind of license it is:

            		if (line.IndexOf("9002") == 5)
                        {
                            RetVal = LICENSE_VERSION.STANDARD;
                        }
                        if (line.IndexOf("13529") == 5)
                        {
                            RetVal = LICENSE_VERSION.PRO;
                        }
                        if (line.IndexOf("723") == 5)
                        {
                            RetVal = LICENSE_VERSION.FREE;
                        }