## ListObjectの使い方 (Excel VBA)

- `.HeaderRowRange`
  + `.DataBodyRange`が存在しない場合、`.HeaderRowRange`からアドレスが取得できる。
- `.DataBodyRange`
  + `Cells`と同じような使い方  
  → `.DataBodyRange(1, 1)`
  + テーブルのデータ領域を2次元配列で格納

    ```
    Dim tableData As Variant
    tableData = .DataBodyRange
    ```
    
  + 2次元配列データをテーブルのデータ領域に貼り付け
    ```
    .DataBodyRange.Resize(Ubound(tableData), Ubound(tableData, 2)) = tableData
    ```
  + データエリアの削除  
  → `.DataBodyRange.Delete`
  + データエリアの存在確認のIf文  
    ```
    If .DataBodyRange Is Nothing Then     '← データがない場合
    If Not .DataBodyRange Is Nothing Then '← データがある場合
    If .DataBodyRange Is Not Nothing Then '← Notの位置が不正なためエラーになる
    ```

- `.ListRows`
  + `.ListRows(i)`のiは1から。 → `For i = 1 To .ListRows.Count`
  + `.ListRows.Count` → テーブル内のデータの行数  
  (`.DataBodyRange.Rows.Count`と同等だが、データ行が無い場合、`.DataBodyRange`は`Nothing`となるため、`.DataBodyRange.Rows.Count`ではエラーになる。)
  + `.ListRows.Add` → テーブル末尾にデータ行の追加
  + `.ListRow(i).Delete` → データ行の削除
  + `.ListRows(i).Range(j)` → .DataBodyRange(i, j)と同じ
- .ListColumns
  + `.ListColumns(i).Range` → テーブルのi列目のデータ (見出し行も含む)
  + `.ListColumns(i).DataBodyRange` → テーブルのi列目のデータ (データ行のみ)
  + `.ListColumns(i).Name` → テーブルのi列目のフィールド名

* * *

- テーブルのフィルタの設定解除
  + `.ShowAutoFilter = False: .ShowAutoFilter = True`
- 書式がおかしくならないように、リストの1行目だけ残して削除する方法 (1行目の書式を残す)
    ```
    Dim i As Long
    For i = 2 To .ListRows.Count: .ListRows(i).Delete: Next
    .ListRows(1).Range.ClearContents
    ```

