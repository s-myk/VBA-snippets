## Sort メソッド, Sort オブジェクト on ListObject

1. Sort メソッド  
   https://excelwork.info/excel/cellsortmethod/

    ```
    .DataBodyRange.Sort _
        key1:= .ListColumns(1).Range, order1:=xlAscending
    ```

1. Sort オブジェクト  
   https://excelwork.info/excel/cellsortcollection/

    ```
    With .Sort
        With .SortFields
            .Clear
            .Add Key:=.Parent.Parent.ListColumns(3).Range, _
                SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        End With
        .MatchCase = False
        .SortMethod = xlStroke
        .Apply
    End With
    ```
