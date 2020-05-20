## VBA ListObjectの使い方

- `.HeaderRowRange`
- `.DataBodyRange`

    ```
    tableData = .DataBodyRange '→ テーブルのデータ領域を2次元配列で格納
    ```
- `.ListRows`
  + `.ListRows.Count` → テーブル内のデータの行数  
  (`.DataBodyRange.Rows.Count`と同等だが、データ行が無い場合、`.DataBodyRange`は`Nothing`となるため、`.DataBodyRange.Rows.Count`ではエラーになる。)
  + `.ListRows.Add` → テーブル末尾にデータ行の追加
  + `.ListRow(i).Delete` → データ行の削除
  + `.ListRows(i).Range(j)` → .DataBodyRange(i, j)と同じ
- .ListColumns
  + `.ListColumns(i).Range` → テーブルのi列目のデータ (見出し行も含む)
  + `.ListColumns(i).DataBodyRange` → テーブルのi列目のデータ (データ行のみ)
  + `.ListColumns(i).Name` → テーブルのi列目のフィールド名

テーブルのフィルタの設定解除 (「Criteria1:=～」を指定していないためフィルタの設定が解除される。)
```
For i = 1 To .ListColumns.Count
    .Range.AutoFilter Field:=i
Next
```
