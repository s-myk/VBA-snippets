# Sort メソッド, Sort オブジェクト

## ListObject に対する Sort

1. Sort メソッド  
   https://excelwork.info/excel/cellsortmethod/

   ```vb
   .DataBodyRange.Sort _
       key1:=.ListColumns(1).Range, order1:=xlAscending, _
       Header:=xlYes
   ```

1. Sort オブジェクト  
   https://excelwork.info/excel/cellsortcollection/

   ```vb
   With .Sort
       With .SortFields
           .Clear
           .Add Key:=.Parent.Parent.ListColumns(1).Range, _
               SortOn:=xlSortOnValues, _
               Order:=xlAscending, _
               DataOption:=xlSortNormal
       End With
       .Header = xlYes
       .MatchCase = False
       .SortMethod = xlStroke
       .Apply
   End With
   ```
