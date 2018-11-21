# HR Management Tool _ Identifying changes in yearly worksheets within the Employee Database

The aim of this code is to compare and identify the changes between two worksheets that contain yearly employee data. We focused on what we deemed as the most important changes (variables) contained in a company’s employee data set. These variables are: employee id (primary key), name, designation of employee, salary and location. The code aims at identifying the new, ex-employees and the employees that stayed in the company. Additionally, the code is designed to display qualitative (e.g. changes in employee designation) as well as quantitative differences (salary changes). This will serve as a useful tool for HR managers (end-users).
For the comparison, we are using one workbook (“HR Management Tool”) with two worksheets “Employees2017”, “Employees2018”. The results from each comparison are displayed in a separate Worksheet (“New Employees”, “Ex-Employees”, “Existing Employees”, “Salary Difference”, “Designation Difference”). 
Final Product: the end-product of this code is a user form with 5 main buttons. Each button directs users to a specific, new worksheet.  We opted for a user form as it allows tighter control. The user form was designed to be simple with few controls that are self-explanatory. The 5 main buttons of the user form are described in detail:
	Button “Existing Employees” – shows the corresponding info for the employees that have stayed in the company (Employee Id present in both “Employees2017” and “Employees2018” worksheets) from the updated table (latest year – 2018). The new worksheet contains the update information, from year 2018. 
•	Button “Ex-Employees” – displays info of the employees that left the company (i.e. whose Employee Id is included in the first (“Employees2017”) but not in the last worksheet (“Employees2018”).
•	Button “New Employees” – displays the table info of the employees that were hired (whose Employee Id is only included in the last datasheet (“Employees2018”).
•	Button “Salary Change of Employees”– displays the table info of the employees whose salaries increased or decreased. For that we used, conditional formatting so that promotions are shown in green color and demotions in red.
•	Button for “Designation Changes” – displays the info of the employees whose designation changed in the last 2 years. The new designation is highlighted in yellow.

To facilitate the end-users, a print button is included next to each main button of the user form to print the corresponding information in pdf. Finally, we included a clear button in each of the new worksheets to clear the contents and display the information of interest anew. 
Process: 
	Identifying new, ex- or existing employees

1)	We defined 3 arrays 2 containing the employee ids (primary keys based on which we searched the original worksheets) and 1 to store the values we are looking for (matches/mismatches). We also specified the dimensions of the arrays. 

2)	We then looped through the employee id arrays for matches (for the “Existing Employees”) and mismatches (2 types: employee ids existing only in the first but not the second worksheet (“Ex-Employees”) and vice versa (“New Employees”)) between the two worksheets (“Employees2017”, “Employees2018”). We iterated by index.

3)	The resulting values were stored in a new array. A for each loop was then used to iterate over all the elements of the array. In this way, only the info of the specific group of interest was retrieved from the original worksheet and was displayed in a new named worksheet (“New Employees, “Ex-Employees”, “Existing Employees”).

	Promotion & Designation Changes

For the promotion and designation change worksheets, a similar approach was implemented. A for loop was used to search through the elements of the array containing the matched employees ids and the resulting info was displayed in new worksheets. We used conditional formatting so that promotions are displayed in green and demotions in red. With regard to the “Designation Change” worksheet, in case of a designation change the new positions are highlighted in yellow.

As mentioned above, all the resulting information from the comparisons is displayed in new worksheets. This will facilitate the process of an HR Manager and will allow him/her to efficiently perform comparisons and calculations, if needed. An additional operation was included to easily print the pdf reports with the employee data based on the query of interest. These files are saved in: 
C :\ Users\ Public
Optimisation Techniques :
Macro-Optimisation Techniques:
1)	Creation of New Index 2D arrays (“arrListEmp17”, “arrListEmp18”) to search the original arrays (“arrEmpExist”, “arrEmpLeft”, “arrEmpNew”) in order to increase efficiency. 
Micro-Optimisation Techniques: 
1)	Early-binding: Declared variables (e.g “arrListEmp18”, “arrListEmp18”) of specific data type (variant) and set them to the named sheet instead of repeatedly referring to the named sheets.
2)	Use of Variant Arrays: for faster processing we stored and processed the range of cells in variant array .
3)	We iterated the VBA arrays by index for faster performance (by using for next)

 
4)	We used multiple if statements to check about the matches/mismatches rather than using Select Case Constructs, as seen in the example above. 
5)	If bmatch is True was removed and instead we used an If not … then statement to limit the processing cycles.
6)	Turned off Screen Updating and Automatic Calculation in the beginning of the sub procedure. 
7)	We did not use select arguments to increase code efficiency.
8)	Use of matching data types to perform operations. For instance, in the if construct in the example above, matching data types (arrays) are compared with each other:  arrListEmp17 is compared to arrListEmp18. 

Comments from Discussion: From the Discussion in the class, we got several comments regarding the complexity of our approach. Initially, we intended to compare many differences including changes in datatypes between worksheets. However, following our group discussion we decided to focus on the changes we thought as the most important for an HR Manager to track. We thus reduced the number of changes we wanted to identify, and focused our code and project on identifying the key differences between worksheets – that is changes in employee status (new employee, ex, existing), designation and salary changes.  Our final product is a form that is simple and easy to use. 
References: https://stackoverflow.com/questions/28684942/appending-a-dynamic-array-in-vba
https://stackoverflow.com/questions/35142985/vba-change-color-of-cells-based-on-value-in-particular-cell
https://ccm.net/forum/affich-540822-excel-vba-change-cell-color-based-on-value
----
