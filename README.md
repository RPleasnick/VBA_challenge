In this challenge, there is a years worth of information for different stocks.  The individual stocks were summarized for the year and then compared.  This is done through VBA in an Excel spreadsheet.
The first piece of code is called stocks.bas and is located:
https://github.com/RPleasnick/VBA_challenge/blob/master/stocks.bas
This subroutine called sheets setup headers and formats the columns.  It then summerizes the individual stocks for the year.  It calculates the yearly change, the percentage change, and the total volume.  This is then into a list.  The program also keeps tract of the stocks that had the greatest percent increase and decrease, and the greatest total volume which gets displayed.
I used "Xpert" to help with the cell formatting.

The second piece of code is called sheets.bas and is located:
https://github.com/RPleasnick/VBA_challenge/blob/master/sheets.bas
This subroutine was taken from "Xpert" and is used to cycle through all the sheets in the workbook.  When run, it calls the sheets subroutine for each sheet in the notebook.

Please see the screenshots of the results.
