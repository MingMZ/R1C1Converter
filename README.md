# R1C1 Converter

R1C1 converter provide easy methods of converting from A1 style to R1C1 style reference used in Excel spreadsheet document.

## Example

Convert from A1 to R1C1:
```csharp
// convert reference "A2" to row and column number
R1C1Converter.R1C1Converter.ToR1C1("A2", out int r, out int c);
// r = 2, c = 1
```

Convert from R1C1 back to A1:
```csharp
// convert row 2 and column 1 back to A1 reference
string a = R1C1Converter.R1C1Converter.FromR1C1(2, 1);
// a = "A1"
```
