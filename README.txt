This application automates the application process for Chicago Department of Transportation's Periodic DAS-Distributed Antenna System permits.

        From the console, this application will:
          1.Request user information
            -name
            -phone
            -email
            -CDOT username
            -CDOT password
          2.Read a file called "excel_files/read/Applications_To_Apply_For.xlsx"
            -Determine if there are equal number of Project Names and CBD ID's
          3.Use list of Project Names and CBD ID's to navigate the CDOT website and apply for the permit.
            -Update "excel_files/read/Applications_To_Apply_For.xlsx" with:
                1.CDOT website's generated CDOT Application Number
          4.Notify when complete. 
          
          
  To date, this application has been used for over 1000 applications to the Chicago Department of Transportation. 
