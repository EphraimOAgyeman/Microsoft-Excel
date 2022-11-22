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
