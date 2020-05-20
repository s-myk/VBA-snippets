## ListObjectの使い方 (Excel VBA)

- `.HeaderRowRange`
  + `.DataBodyRange`が存在しない場合、`.HeaderRowRange`からアドレスが取得できる。
  + ヘッダー行の表示/非表示: `tbl.ShowHeaders = True/False`
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

- 初期化 (1行目の数式と表示形式を残す)

    ```
    Dim i As Long, 数式 As String, 表示形式 As String
    If .ListRows.Count >= 2 Then
        .DataBodyRange(2, 1).Resize(.ListRows.Count - 1, .ListColumns.Count).Delete
    End If
    If .ListRows.Count > 0 Then
        For i = 1 To .ListColumns.Count
            With .ListRows(1).Range(i)
                数式 = vbNullString
                If .Value <> "=*" Then
                    数式 = .Value
                End If
                表示形式 = .NumberFormatLocal
                .Clear
                .Value = 数式
                .NumberFormatLocal = 表示形式
            End With
        Next
    Else
        .ListRows.Add
        .ListRows(1).Range.Clear
    End If
    ```
