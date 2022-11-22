# Microsoft Excel
All about excel

## Shortcuts
> `ctrl` + `arrow keys` = move to the last column or row in the direction of the arrow keys

> `ctrl` + `shift` + `arrow keys` = selects the area it covers

> `ctrl` + `A` = selects the whole table

> `ctrl` + `PgUp`/`PgDn` = Navigate through sheets in a workfile.

> `ctrl` + `1` = Formatting options pop up
 
## Fill option
This is used to easily fill data

STEPS
> Click on the cell you want to use to fill other cells

> Then Click on `Fill` in the right most part of the `Home` tab

> Fill in the options and Viola


## Hide
You can right click on a table, sheet, row or column and choose hide and unhide when desired.
> Make sure you check for hidden sheets when you recieve a file or when you send a file


## Sizing up the columns
Columns spaces can be different and changing them all can be time consuming.
> Select all the columns at the alphabet names and adjust one column and that adjusment will be sent to the rest of the columns

> Select all the columns and double click in between two columns, and Excel will automatically assign sufficient space to each column 🎊


# Formulas and Functions
Making calculations with the cells.
> Click on an empty cell and type your formula there or on the formula bar
```
#Average = total over count
=(A1+A2+A3)/3
```
> This will give you a value. 

> Copy the cell you wrote the formula in side, select the cells you want to have it pasted inside, and paste

> With that, you will duplicate the formula for other places without typeing it all over again 🎊

---

## Logical Conditions
> Logical symbols are formulas so they start with `=`.
```
= D10 > H50
```

| Logical Symbols | Meaning | 
| ------------- |-------------:| 
| >     | A greater than B | 
| <   | A less than B     |  
| <> | A is not equal to B    |   
| =  | A equal to B | 
| >= | A is greater than or equal to B    |  
| <=| A is less than or equal to B   |  

## Cell references
We learnt earlier that the cell formulas can be copied. Now we want it to be copied intelligently.
By this sometimes we need the column to be fixed or the rows.
> The dollar sign does this `$`

> =`$K10` the dollar being by just the `k` means the column is fixed but the rows can move

> =`k$10` means the column can move but the rows cannot 

> =`$k6*(1-j$3)`meabs column k is fixed and the its rows will move, row 3 is fixed but its columns will move


## Aggregation function
First few ones to know are `SUM`, `AVERAGE`, `COUNT`

```
=SUM(H1:H4)
```
> You can copy the code and paste it into different cells


# Logical functions

## AND/OR
One logical function can be if A is greater than B. -
Another L.Funcion is if C is less than D.

AND / OR check the outcome of both or more functions together.

Quesiton sample - Average booking window `AND`  Average length of stay for hotel 1 `is shorter than` in hotel 2.
```
=AND(logical condition, Logical Condition, ...)
```

## IF
Conditional functions checks for true or false then makes a move.
```
=IF(A>B),RETURN ACTION,ELSE
```

You will know the number of IFs you need by knowing the number of outcomes you are looking for;

- If you are looking for `2` outcomes, its `2 -1 = 1 IF statement`.
- If you have `3` outcomes, its `3 -1 = 2 IF statements`, and so forth

## Nested IF

You can insert the second or third IF statement at where the else parameter is placed.
```
=IF(Condition), Return Action, IF(Condition 2), Return Action, Else Return Action
```

> NB: The conditions can have AND/OR and be sure to add indenting to code for readability and more

> Put statement returns in quotes

## IFS

`IFS` is a replacement to the nested if statement
```
=IFS(Condition, Return Value,Condition, Return Value,Condition, Return Value)
```

##  Aggregation functions + IF

How many people paid more than 400
- COUNTIF - =COUNTIF(range[where to count], Condition for if [in quotes])

How any people payed above 400 in 2022
- COUNTIFS [count plusmultiple If statements ] - =COUNTIFS(range1, Criterial/condition1, range2, Criteria2/condition2 ...)
