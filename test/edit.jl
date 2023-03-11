import XLSX
import Tables
using Test, Dates
import DataFrames

dir = raw"C:\work\XLSX"

filename = raw"test.xlsx"

xf = XLSX.open_xlsx_template(joinpath(dir, filename)) # openxlsx does not allow saving
XLSX.writexlsx(joinpath(dir, "test_simply_saved.xlsx"), xf; overwrite=true)

# TODO: Row formatting is lost (green line) --> OK
# TODO: Formulae that are only referenced are lost --> OK; but what if someone overwrites the original formula?

xf["Values"]["B1"] = "Just replacing a string"
xf["Formatting"]["B2"] = "Still red" # Preserves formatting correctly
xf["Formatting"]["B2"] = 0 # CHange of type (from string to Int) --> destroys formatting
XLSX.setdata!(xf["Formula"], "A1", [2,3,4,5], 1) # Same type --> formula must be re-evaluated with Ctrl+Alt+Shift+F9
XLSX.writexlsx(joinpath(dir, "test_modified.xlsx"), xf; overwrite=true)
