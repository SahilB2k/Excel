## **V Lookup = Vertical lookup**



main formula: **=VLOOKUP(CellWhereDataIsEntered,CompleteDataTable,ParticularColumnIndex,RangeLookUp)**

		eg - =VLOOKUP($H$2,$A$1:$E$13,ROW()-1,FALSE)

***(NOTE - for lookup to be dynamic select the "*CellWhereDataIsEntered,CompleteDataTable" <i>and click f4 to add dollar between the values so that the data should change dynamically)</i>**

**<i> without the dollar we cannot flash fill the records because they are relative reference based and mis match error happens</i>**



for ease of use instead of selecting the completing table for adding on into the vlookup formula name the table and this name can be used in the formula of vlookup



* how to name a complete table

&nbsp;	1.Select the required table 

&nbsp;	2.On the top right bellow the ribbon there would be dropdown named as **NAMEBOX** , in that dropdown write the 	name of the table accordingly(do no start the name with a number or 	space)

&nbsp;	3.press enter and your table will be named as mentioned

&nbsp;	4.Once you name a table from the namebox you cannot then change or delete from the namebox

&nbsp;	5.for renaming or deleting go in formulas section and search for name manager

&nbsp;	6 in the name manager you can rename or delete the name of the table



(*note : in vlookup we have **range lookup** in which there are two options 1.Approximation match (0),2.exact match(1) , for approximation match it is compulsory that the data should be in ascending order , if not error will be raised)*



*exact match are used when there are fixed values like name , age*

*approximation match is used when their are range of values like salary*



* <b>when and why $ is used:</b>

	1.when after every change if we want to fetch vale from a single column/cell of the table    	without getting iterated(eg H1 -> H2 -> H3...) we use $ to fix the column/cell **this is 	called absolute reference**

&nbsp;	2. when after every change in column the data referred or fetched should be changed 	according to the column number of name then no $ sign in used **this is called relative 	reference**



* **Match function:**  used to define the column number from a particular column

&nbsp;	Syntax =MATCH(theColumn/cellwhichIsToBeMatchedWith,AllColumnsOfTheTable,ExactMatch(0))

&nbsp;	eample:  =VLOOKUP($I$3,$A$1:$E$13,**MATCH(H4,$A$1:$E$1,0)**,0)



* **INDEX Function :** used to provide a particular x(row) and Y(column) matching value from the 	table example- fetching the teacher name for a subject(x) at a particular time(y)

&nbsp;	Syntax: =**Index( TableData , RowNumberOfXCoordinate,ColumnNumberOfYCoordinate)**

&nbsp;	example- **=INDEX(A1:J9,C17,C18)**



&nbsp;	**(*NOTE: for Index function to use first use match function to find the row and column 	number of the data required and then use this values in the index formula)***



* <b>=SUMIF() : </b>adds the cell specified to a given condition(for eg- find total quantity sold 	according to the region i.e - east , west , south, if i select south then it will 	calculate the total quantity on its own for all entries of south and will give me the 	final count. )

&nbsp;	Syntax: =SUMIF(**range**=the categories from which we have to count eg.region ,**criteria** = the region which we mention eg.east , west etc , **sum\_range** = the range of values eg. quantities )














