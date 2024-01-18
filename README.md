--------------------------------------------------------------
-----------      MULTIPLE YEAR STOCK VBA CODE       ----------

-----------             READ ME FILE                ----------

--------------------------------------------------------------

--                    by LAURA VARA                         --

--------------------------------------------------------------
                SECTIONS IN THIS READ ME FILE
---------------------------------------------------------------
I.  - Documents in this repository
II. - Instructions of the challenge
III.- References to the origin of the code

______________________________________________________________
--------------------------------------------------------------
               I. DOCUMENTS IN THIS REPOSITORY
--------------------------------------------------------------
1) Folder with screen shots of:
	• A blank version of VBA Excel file before the code
	• 3 Screen shots. One per each year of stock to loop
	• 6 Screen shots. One per each section of the 
	Developer Module with the code as it scrolls down.

2) The Excel File with Macros enabled to view the VBA code
from the Developer's tab

3) A Visual Studio Code file with the full code for this challenge

4) This ReadMe File
______________________________________________________________
--------------------------------------------------------------
               II. INSTRUCTIONS FOR CHALLENGE
--------------------------------------------------------------

1) Create a script that lopos through all the stocks for one year
and outputs the following information:
	• The ticker symbol.
	• Yearly change from the opening price at the beginning
	of a given year to the closing price  at the end of that 
	year.
	• The percentage change from the opening price at the 
	beginning of a given year to the closing price at the end
	of that year.
	• The total stock volume of the stock.
Note: an image was provided that should match my results. Which
by the end of my code, my file matches the image provided.

2) Add functionality to your script to return the stock with the
	• Greatest % increase
	• Greatest % decrease
	• Greatest total volume
Note: An image was provided that should match my results. Which 
also does match my addition to the code.

3) Make the appropriate adjustments to your VBA script to enable it
to run on every worksheet(the current, and every year included) at once.

4) Make sure to use conditional formatting that will highlight positive 
change in green and negative change in red. 

_________________________________________________________________
-----------------------------------------------------------------
               III. REFERENCES TO MY CODE
-----------------------------------------------------------------
Note: If the code is referenced from a class, it will say "Class"
-----------------------------------------------------------------

• DECLARING THE WORKSHEET/VARIABLES/FINDING THE LAST ROW
All of this section was covered in Class. Mainly in the last two examples,
the Credit Card and the Census examples. This is where the 
"For Each ws In Worksheets" came in, allowing the code to pass to the
next worksheets in this document.
This is also where we learned about finding the LastRow.

• THE LOOP
All of this section was covered in Class. Although it was tricky,
this part of the code came from the Credit Card and Census examples.
Using the ws before 'cells' and 'ranges' allowed me to pass all the 
commands to the next worksheets.
We saw concatanation in class as well, which is shown in the values
of the Summary Table Row. ws.Range("I" & Summary_Table_Row).Value = Ticker
Color formatting was seen in several exercises in class.

**Note that in the loop, you'll find a comment with an additional code
if I were to follow the grading segment  to add ($ Currency) to the Yearly Column.
Note that your initial instructions to follow and match the images, do 
not include a ($ Currency) symbol. So it was done in the comments.
That code was seen in class. It's in the Census code for correcting the
currency format.

• FORMATTING OUTSIDE THE LOOP
**Number Format: 
https://www.educba.com/vba-number-format/
https://www.mrexcel.com/board/threads/custom-number-format-with-special-character-macro.1089618/

**Auto Fit Format:
https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

• GREATEST % INCREASE, DECREASE AND GREATEST TOTAL VOLUME
This is pretty much the same as the above code, thus was seen in class.

• MAX PERCENT INC/DEC/TOTAL VOL LOOP
We saw this in class. There was an exercise called Crypto Kennel pt3
and also referenced with the credit card and census examples. It pretty
much follows similar to the loop for the ticker.

• OUTPUT THE VALUES IN THE 2ND SUMMARY TABLE
Follows the same structure of the ticker Summary Table.

• THE FORMATING
**Number Format: 
https://www.educba.com/vba-number-format/
https://www.mrexcel.com/board/threads/custom-number-format-with-special-character-macro.1089618/

------------------------------------------------------------------------
________________________________________________________________________

