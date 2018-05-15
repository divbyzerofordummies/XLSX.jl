
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

julia> XLSX.getdata("myfile.xlsx", "mysheet", "A2:B4") # shorthand for all above
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

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.gettable("myfile.xlsx", "mysheet")...)
3×2 DataFrames.DataFrame
│ Row │ HeaderA │ HeaderB  │
├─────┼─────────┼──────────┤
│ 1   │ 1       │ "first"  │
│ 2   │ 2       │ "second" │
│ 3   │ 3       │ "third"  │
```

## Streaming Large Excel Files and Caching

The method `XLSX.openxlsx` has a `enable_cache` option to control worksheet cells caching.

Cache is enabled by default, so if you read a worksheet cell twice it will use the cached value instead of reading from disk
in the second time.

If `enable_cache=false`, worksheet cells will always be read from disk.
This is useful when you want to read a spreadsheet that doesn't fit into memory.

The following example shows how you would read worksheet cells, one row at a time,
where `filename.xlsx` is a spreadsheet that doesn't fit into memory.

```julia
julia> f = XLSX.openxlsx("filename.xlsx", enable_cache=false)

julia> sheet = f["sheetname"]

julia for r in XLSX.eachrow(sheet)
          # r is a `SheetRow`. Values are read using column references.
          v1 = r[1]    # will read value at column 1
          v2 = r["B"]  # will read value at column 2
      end
```

You could also stream tabular data using `XLSX.eachtablerow(sheet)`, which is the underlying iterator in `gettable` method.
Check docstrings for `XLSX.eachtablerow` for more advanced options.

```julia
julia> for r in XLSX.eachtablerow(sheet)
           # r is a `TableRow`. Values are read using column labels or numbers.
           v1 = r[1] # will read value at table column 1
           v2 = r[:COL_LABEL2] # will read value at column labeled `:COL_LABEL2`.
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
