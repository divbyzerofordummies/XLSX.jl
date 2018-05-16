
# XLSX.jl

[![License](http://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat)](LICENSE)
[![Build Status](https://travis-ci.org/felipenoris/XLSX.jl.svg?branch=master)](https://travis-ci.org/felipenoris/XLSX.jl)
[![codecov.io](http://codecov.io/github/felipenoris/XLSX.jl/coverage.svg?branch=master)](http://codecov.io/github/felipenoris/XLSX.jl?branch=master)
[![XLSX](http://pkg.julialang.org/badges/XLSX_0.6.svg)](http://pkg.julialang.org/?pkg=XLSX&ver=0.6)

Excel file parser written in pure Julia.

## Installation

```julia
julia> Pkg.add("XLSX")
```

## Basic Usage

The basic usage is to read an Excel file and read values.

```julia
julia> import XLSX

julia> xf = XLSX.openxlsx("myfile.xlsx")
XLSXFile("myfile.xlsx")

julia> XLSX.sheetnames(xf)
3-element Array{String,1}:
 "mysheet"
 "othersheet"
 "named"

julia> sh = xf["mysheet"] # get a reference to a Worksheet
XLSX.Worksheet: "mysheet". Dimension: A1:B4.

julia> sh["B2"] # From a sheet, you can access a cell value
"first"

julia> sh["A2:B4"] # or a cell range
3×2 Array{Any,2}:
 1  "first" 
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", "mysheet", "A2:B4") # shorthand for all above
3×2 Array{Any,2}:
 1  "first" 
 2  "second"
 3  "third"

julia> sh[:] # all data inside worksheet's dimension
4×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"  
 2           "second" 
 3           "third"

julia> xf["mysheet!A2:B4"] # you can also query values from a file reference
3×2 Array{Any,2}:
 1  "first" 
 2  "second"
 3  "third"

julia> xf["NAMED_CELL"] # you can even read named ranges
"B4 is a named cell from sheet \"named\""

julia> xf["mysheet!A:B"] # Column ranges are also supported
4×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
 3           "third"

julia> close(xf) # close the file when done reading
```

To inspect the internal representation of each cell, use the `getcell` or `getcellrange` methods.

The example above used `xf = XLSX.openxlsx(filename)` to open a file, so the contents will be fetched from disk as needed
but you need to close the file when done reading with `close(xf)`.

You can also use `XLSX.readxlsx(filename)` to read the whole file and return a closed `XLSXFile`.

## Read Tabular Data

The `gettable` method returns tabular data from a spreadsheet as a tuple `(data, column_labels)`.
You can use it to create a `DataFrame` from [DataFrames.jl](https://github.com/JuliaData/DataFrames.jl).
Check the docstring for `gettable` method for more advanced options.

There's also a helper method `readtable` to read from file directly, as shown in the following example.

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")...)
3×2 DataFrames.DataFrame
│ Row │ HeaderA │ HeaderB  │
├─────┼─────────┼──────────┤
│ 1   │ 1       │ "first"  │
│ 2   │ 2       │ "second" │
│ 3   │ 3       │ "third"  │
```

## Writing Excel Files

You can export tabular data to Excel using `XLSX.writetable` method.

```julia
julia> import DataFrames, XLSX

julia> df = DataFrames.DataFrame(:integers=>[1, 2, 3, 4], :strings=>["Hey", "You", "Out", "There"], :floats=>[10.2, 20.3, 30.4, 40.5], :dates=>[Date(2018,2,20), Date(2018,2,21), Date(2018,2,22), Date(2018,2,23)], :times=>[Dates.Time(19,10), Dates.Time(19,20), Dates.Time(19,30), Dates.Time(19,40)], :datetimes=>[Dates.DateTime(2018,5,20,19,10), Dates.DateTime(2018,5,20,19,20), Dates.DateTime(2018,5,20,19,30), Dates.DateTime(2018,5,20,19,40)]):floats=>[10.2, 20.3, 30.4, 40.5])
4×6 DataFrames.DataFrame
│ Row │ integers │ strings │ floats │ dates      │ times    │ datetimes           │
├─────┼──────────┼─────────┼────────┼────────────┼──────────┼─────────────────────┤
│ 1   │ 1        │ Hey     │ 10.2   │ 2018-02-20 │ 19:10:00 │ 2018-05-20T19:10:00 │
│ 2   │ 2        │ You     │ 20.3   │ 2018-02-21 │ 19:20:00 │ 2018-05-20T19:20:00 │
│ 3   │ 3        │ Out     │ 30.4   │ 2018-02-22 │ 19:30:00 │ 2018-05-20T19:30:00 │
│ 4   │ 4        │ There   │ 40.5   │ 2018-02-23 │ 19:40:00 │ 2018-05-20T19:40:00 │

julia> XLSX.writetable("df.xlsx", DataFrames.columns(df), DataFrames.names(df))
```

## Streaming Large Excel Files and Caching

The method `XLSX.openxlsx` has a `enable_cache` option to control worksheet cells caching.

Cache is enabled by default, so if you read a worksheet cell twice it will use the cached value instead of reading from disk
in the second time.

If `enable_cache=false`, worksheet cells will always be read from disk.
This is useful when you want to read a spreadsheet that doesn't fit into memory.

The following example shows how you would read worksheet cells, one row at a time,
where `myfile.xlsx` is a spreadsheet that doesn't fit into memory.

```julia
julia> f = XLSX.openxlsx("myfile.xlsx", enable_cache=false)

julia> sheet = f["mysheet"]

julia> for r in XLSX.eachrow(sheet)
          # r is a `SheetRow`, values are read using column references
          rn = XLSX.row_number(r) # `SheetRow` row number
          v1 = r[1]    # will read value at column 1
          v2 = r["B"]  # will read value at column 2
       end
```

You could also stream tabular data using `XLSX.eachtablerow(sheet)`, which is the underlying iterator in `gettable` method.
Check docstrings for `XLSX.eachtablerow` for more advanced options.

```julia
julia> for r in XLSX.eachtablerow(sheet)
           # r is a `TableRow`, values are read using column labels or numbers
           rn = XLSX.row_number(r) # `TableRow` row number
           v1 = r[1] # will read value at table column 1
           v2 = r[:HeaderB] # will read value at column labeled `:HeaderB`
       end
```

## References

* [ECMA Open XML White Paper](https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf)

* [ECMA-376](https://www.ecma-international.org/publications/standards/Ecma-376.htm)

* [Excel file limits](https://support.office.com/en-gb/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)

## Alternative Packages

* [ExcelFiles.jl](https://github.com/davidanthoff/ExcelFiles.jl)

* [ExcelReaders.jl](https://github.com/davidanthoff/ExcelReaders.jl)

* [XLSXReader.jl](https://github.com/mpastell/XLSXReader.jl)

* [Taro.jl](https://github.com/aviks/Taro.jl)
