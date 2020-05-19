## AutoFilterの使い方

状態の確認と初期化：
1. Worksheetとそこに含まれるListObjectに対しての操作
   - 絞り込みの有無: `ws.FilterMode`
   - 絞り込みの解除: `ws.ShowAllData` (.FilterMode = Falseの時に実行するとエラー)
2. Worksheet
   - オートフィルターの有無: `ws.AutoFilterMode`
   - オートフィルターの絞り込みの有無: `ws.AutoFilter.FilterMode` (*1)
   - オートフィルターが設定されている範囲: `ws.AutoFilter.Range` (*1)
   - オートフィルターの絞り込みの解除: `ws.AutoFilter.ShowAll` (*2)
   - オートフィルターの削除1: `ws.AutoFilterMode = False` (*1)
   - オートフィルターの削除2: `ws.AutoFilter.Range.AutoFilter` (*1)
   - (*1 `ws.AutoFilterMode = False`の時に実行するとエラー)   
   (*2 `ws.AutoFilter.FilterMode = False`の時に実行するとエラー)
3. ListObject
   - オートフィルターの有無: `lo.ShowAutoFilter`
   - オートフィルターの絞り込みの有無: `lo.AutoFilter.FilterMode` (*3)
   - オートフィルターの絞り込みの解除1: `lo.ShowAutoFilter = False` → True
   - オートフィルターの絞り込みの解除2: `lo.AutoFilter.ShowAll` (*4)
   - (*3 `lo.ShowAutoFilter = False`の時に実行するとエラー)   
   (*4 `lo.AutoFilter.FilterMode = False`の時に実行するとエラー)
   - オートフィルターの絞り込みの解除3: `For i = 1 To .ListColumns.Count: .Range.AutoFilter Field:=i: Next` (*3)
   - オートフィルターの解除: `lo.ShowAutoFilter = False` (TrueとFalseでフィルタボタンの表示切替)
* * *
フィルタのOn/Off (ListObjectでのサンプル)：
```
If .AutoFilter.FilterMode Then
    .AutoFilter.ShowAllData
Else
    .Range.AutoFilter Field:=1, Criteria1:=Array("="), Operator:=xlFilterValues, _
        Criteria2:=Array(2, "1926/12/25", 2, "1989/1/8", 2, "2019/5/1")
End If
```
* * *
`.AutoFilter Field:=iCol, Criteria1:="<4/30/2019"` ‘以前  
`.AutoFilter Field:=iCol, Criteria1:=">=1/8/1989"` ‘以降

xlAnd / xlOr / xlFilterValues
 
参考資料：
- The Ultimate Guide to Excel Filters with VBA Macros – AutoFilter Method  
https://www.excelcampus.com/vba/macros-filters-autofilter-method/
- How to Filter for Dates with VBA Macros in Excel  
https://www.excelcampus.com/vba/filter-dates/
- オートフィルタの設定と解除  
https://vbabeginner.net/%E3%82%AA%E3%83%BC%E3%83%88%E3%83%95%E3%82%A3%E3%83%AB%E3%82%BF%E3%81%AE%E8%A8%AD%E5%AE%9A%E3%81%A8%E8%A7%A3%E9%99%A4/
- オートフィルタ(AutoFilter)でのデータ抽出  
http://www.eurus.dti.ne.jp/~yoneyama/Excel/vba/vba_autofilter.html
