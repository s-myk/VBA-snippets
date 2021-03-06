# AutoFilter の使い方

## 状態の確認と初期化

1. Worksheet とそこに含まれる ListObject に対しての操作

   - 絞り込みの有無: `ws.FilterMode`
   - 絞り込みの解除: `ws.ShowAllData` (.FilterMode = False の時に実行するとエラー)

2. Worksheet

   - オートフィルターの有無: `ws.AutoFilterMode`
   - オートフィルターの絞り込みの有無: `ws.AutoFilter.FilterMode` (\*1)
   - オートフィルターが設定されている範囲: `ws.AutoFilter.Range` (\*1)
   - オートフィルターの絞り込みの解除: `ws.AutoFilter.ShowAll` (\*2)
   - オートフィルターの削除 1: `ws.AutoFilterMode = False` (\*1)
   - オートフィルターの削除 2: `ws.AutoFilter.Range.AutoFilter` (\*1)
   - (*1 `ws.AutoFilterMode = False`の時に参照/実行するとエラー)  
     (*2 `ws.AutoFilter.FilterMode = False`の時に実行するとエラー)

3. ListObject
   - オートフィルターの有無: `lo.ShowAutoFilter`
   - オートフィルターの絞り込みの有無: `lo.AutoFilter.FilterMode` (\*3)
   - オートフィルターの絞り込みの解除 1: `lo.ShowAutoFilter = False` → True
   - オートフィルターの絞り込みの解除 2: `lo.AutoFilter.ShowAll` (\*4)
   - オートフィルターの絞り込みの解除 3: `For i = 1 To .ListColumns.Count: .Range.AutoFilter Field:=i: Next` (\*3)
   - オートフィルターの解除: `lo.ShowAutoFilter = False` (True と False でフィルタボタンの表示切替)
   - (*3 `lo.ShowAutoFilter = False`の時に参照するとエラー)  
     (*4 `lo.AutoFilter.FilterMode = False`の時に実行するとエラー)

## フィルタの On/Off (ListObject でのサンプル)

- 複数の条件に合致するフィルタ (xlFilterValues を使用)

  ```vb
  If .ShowAutoFilter = False Then
      ' オートフィルタが存在しない場合は何もしない
      ' do Nothing
  ElseIf .AutoFilter.FilterMode Then
      ' すでにオートフィルタでの絞り込みが行われている場合は、絞り込みを解除
      .AutoFilter.ShowAllData
  Else
      ' 絞り込みが行われていない場合は、複数の条件を指定して絞り込み
      .Range.AutoFilter Field:=1, Criteria1:=Array("="), Operator:=xlFilterValues, _
          Criteria2:=Array(2, "1926/12/25", 2, "1989/1/8", 2, "2019/5/1")
  End If
  ```

- 1 つの条件に合致するフィルタ

  ```vb
  '以前
  .AutoFilter Field:=iCol, Criteria1:="<4/30/2019"

  '以降
  .AutoFilter Field:=iCol, Criteria1:=">=1/8/1989"
  ```

- xlAnd / xlOr / xlFilterValues などの使い方  
  https://excel-ubara.com/excelvba1/EXCELVBA389.html

## 参考資料

- The Ultimate Guide to Excel Filters with VBA Macros – AutoFilter Method  
  https://www.excelcampus.com/vba/macros-filters-autofilter-method/
- How to Filter for Dates with VBA Macros in Excel  
  https://www.excelcampus.com/vba/filter-dates/
- オートフィルタの設定と解除  
  https://vbabeginner.net/%E3%82%AA%E3%83%BC%E3%83%88%E3%83%95%E3%82%A3%E3%83%AB%E3%82%BF%E3%81%AE%E8%A8%AD%E5%AE%9A%E3%81%A8%E8%A7%A3%E9%99%A4/
- オートフィルタ(AutoFilter)でのデータ抽出  
  http://www.eurus.dti.ne.jp/~yoneyama/Excel/vba/vba_autofilter.html
