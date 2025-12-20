* changing the format of the word to correct format

 	**=proper(column)**

* changing the words to upppwercase or lowercase:
  **=uppercase(column) , =lowercase(column)**
* removing duplicates:

 	1.select the complete table **(cntrl+shift+>+<)**

 	2.go in data tab in the ribbon and select the **remove duplicate** button and select 	accordingly

* highlighting a complete row based on category:

 	1.select the complete table\*\*(cntrl+shirt+>+<)\*\*

 	2.go to conditional formatting in the ribbon

 	3.select **new rule** option

 	4.select one of the category in the category column , **remove the dollar sign** after the 	number and then after the number **type =** then **add a empty block** in the excel so that the 	table will highlight according to the category you will type into the block(eg-**=$C2=$J$2**)

 	5.select a format(fill color)

* highlighting rows :

 	1.select the complete table\*\*(cntrl+shift+>+downarrow)\*\*

 	2.go to conditional formatting and select highlighting cell rule and select the required 	rule

* deleting empty rows:

 	1.click **cntrl+G(goto)**

 	2.select special

 	3.select blank , then all the blank column/rows would e select

 	4.click **(cntrl)+(-)**  and select the required option from the dialogue block appeared

* removing extra space from the word leading to extra word length

 	1.use **trim** function

 	2.**=substitute(column,char(160),"")**: this would also work sometimes

* **splitting(one of the method)** : two column in one cell (eg-sahil-18,yati-16)

 	1.select and copy the age column

 	2.paste it in the new column and press **cntrl+E(flash filling)**

* **splitting(dynamic method best):** two columns in one cell spearated by a seperator:

 	`1. use **=left(columnname,find("seperator",columnname,1)-1)**

* finding a particular alphabet or symbo from a cell:

 	1.use **=find("symboltofind",columnname,1)**

* **Data Validation:** to make sure that the data entered is in right format(eg phone number should be of 10 digits only)

 	1.select the required column

 	2.Go on **Data section** and select data validation

 	3.select the required validation in the **allow dropdown**

 	3.Go in the **error alert** section and type the required message to be shown to the user

* conditional data entry:

 	use **=if(columnname operator , iftrueprint , itfalseprint) (eg: =IF(D2>=32,"Pass","Fail")**

* Nested if for conditioning(not reliable when large if else conditioning):

 	eg( **=IF(D2>75000,D2\*20%,IF(D2>50000,D2\*15%,D2\*10%))** )

* IFS: for long nested if conditioning , similar to normal IF just no need to write the if word again and again:

 	eg- ( **=IF(D2>75000,D2\*20%,D2>50000,D2\*15%,D2<50000,D2\*10%) )**

* Conditional **AND/OR:**
* 

** **	syntax: **=IF(AND(columnname function value, columnname function 	value),"iftrueprint","iffalseprint")** --same  for OR only instead of AND put OR

 	eg: ( **=IF(AND(D2="Finance",F2>=5),"Yes","No")**  )

* sorting: sorting one or multiple column:
  1.select the complete table

 	2.go in data section and select sort adjust the columns accordingly , also add new column 	if necessary

* Filtering columns:

 	1.use formula - 	**=FILTER(selectCompleteTable,selectParticularColumnWhichNeedToBeFiltered="value","IfNotFou	ndErrorMsg"  (eg :-  =FILTER(A2:H1001,D2:D1001,"Not Found)**

* creating a dropdown for unique elements of a column (eg : state names)

 	1.make a copy of the required column

 	2.select the complete copied column

 	3.remove the duplicate from remove duplicate option from the data section

 	4.now select a cel and click on data validation

 	5.select the list option inside the data validation interface

 	6.in the data validation section select the copied column for entries inside the dropdown 	eg-(=$D:$D)

* mixed functions

 	eg- =SORT(FILTER(A2:I1001,C2:C1001=K2,"NotFound"),9,-1)

 	explaination =SORT(FILTER(culmnname:Alltablevalues,COlumnNamewhichistobefiltered=uniqueValues,"ifNotFoundMsg"),ColumnNumberWhichisToBeSorted,AES(1)/DESC(-1)





----------------------------------------------------------------------------

### Statistical Formulas

---------------------------------------------------------------------------



1\.**=countif(range,criteria)** : used to count a number of specified criteria from a range of values

eg: finding number of total number of object purchased from a range of particular shops individually

syntax=countif(range=selectTheCompleteColumnWhichContainsNumericValues,criteria=SelectTheCompleteColumnWhichConatainTheUniqueCriteria(eg-ShopNames))



2.=countifs(range,criteria\_range1,criteria,criteria\_range2,criteria2) : similar to countif , but allows multiple criteria to be selected



3.=sumif()=adds the cell specified by a given criteria

syntax: sumif(range , criteria , sum\_Range): range= theCompleteColumnOfTheCriteria, criteria=The cell where you write a particular criteria name to get the value , sum\_range=the column where the numeric vales are



4.sumifs() : similar to summif() but only provide multiple criteria selection



5.=AVERAGEIF(range , criteria,average\_range) : finds average of the cells specified by a given criteria



syntax: =AVERAGEIF( range=column which contains all the criteria , criteria a unique criteria from the range of criteria for which you want to count the average value , average\_Range= numeric value which you want to find specifically according to the criterias )



6.=AVERAGEIFS() : similar to averageif() just provides multiple criteria options



7.MAXIF(max\_Range,criteria\_Range1,criteria1): returns largest value in the set of value according to given criteria

syntax= MAXIFS(max\_Range = range of numeric value column from where you want to get the max value, criteria\_range1=column where the range of criterias (eg- different city names) are mentioned , criteria = a specific criteria from the list of criteria coumn for which you want to find the max value)











 



 







 

 

