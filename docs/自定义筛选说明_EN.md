Once the Excel file is open, enter the filter in the filter input box under the 'Operations' tab according to the criteria
## Principle
Use the Python 'eval() function to compute the equation

See [excel.py](https://github.com/Loser123zbx/ZjangDataViewer/blob/master/src/excel.py) for details.

# Fill in the custom guidelines

`x` indicates all values within the item

## Supported symbols

### Logical operations

Equal to `=`

Greater than or equal to `>=`

Less than or equal to `<=`

Greater than `>`

Less than `<`

Not equal to `!=`

### Math

Add ` `

minus `-`

Multiply `*`

Except for `/`

Power `**`

Divisibility `//`

Remainder `%`

### Logical operation: (spaces before and after)

and `and`

or `or`

Non-`not`

### Syntax

Outputs values greater than 114 in the selected column and their corresponding names
> x > 114

Outputs values greater than 114 and their corresponding names in the selected column, and outputs a difference from 114
> x > 114 # x - 114

Outputs values greater than 114 and less than 514 in the selected column and their corresponding names
> x > 114 and x < 514

