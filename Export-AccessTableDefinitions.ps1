# <#
# .SYNOPSIS
# Access データベース (.accdb) のスキーマ情報をテキストファイルに出力します。
# .DESCRIPTION
# 指定されたフォルダ内のすべての .accdb ファイルを検索し、それぞれのデータベースについて、
# テーブル定義、リンクテーブル情報、リレーションシップ、クエリ定義を抽出し、
# "SchemaOutput" フォルダ内にファイル名ごとのスキーマファイル (.txt) を作成します。
#
# 依存関係: Microsoft Data Access Objects (DAO) が必要です。通常、Microsoft Office/Access
#          がインストールされていれば利用可能です。
# #>
Write-Host "Access スキーマ出力スクリプトを開始します..."

$targetFolder = $PSScriptRoot
$outputFolder = Join-Path $PSScriptRoot "SchemaOutput"

# DAO オブジェクトを作成 (Access 2007 以降の .accdb 形式に対応)
try {
    $dao = New-Object -ComObject DAO.DBEngine.120
    Write-Host "INFO: DAO DBEngine Version $($dao.Version) を使用します。"
}
catch {
    Write-Error "DAO オブジェクト (DAO.DBEngine.120) の作成に失敗しました。Microsoft Access または Access Database Engine がインストールされているか確認してください。"
    exit 1
}


# Access データ型と一般的な名称のマッピング (より詳細に)
$typeMap = @{
    1   = "Boolean"; # Yes/No
    2   = "Byte"; # 数値型 (バイト型)
    3   = "Integer"; # 数値型 (整数型)
    4   = "Long"; # 数値型 (長整数型)
    5   = "Currency"; # 通貨型
    6   = "Single"; # 数値型 (単精度浮動小数点数型)
    7   = "Double"; # 数値型 (倍精度浮動小数点数型)
    8   = "Date/Time"; # 日付/時刻型
    10  = "Short Text"; # 短いテキスト (以前の Text)
    11  = "OLE Object"; # OLE オブジェクト
    12  = "Long Text"; # 長いテキスト (以前の Memo)
    15  = "Replication ID"; # レプリケーション ID (GUID)
    16  = "BigInt"; # 大きい数値 (Access 2016+)
    17  = "Binary"; # バイナリデータ (VarBinary など)
    18  = "Date/Time Extended"; # 日付/時刻拡張型 (Access 2016+)
    20  = "Decimal"; # 数値型 (十進型) - DAOではNumericとも
    101 = "Attachment"; # 添付ファイル
    109 = "Multi-valued"; # 複数値を持つフィールド (Complex Textなど) - Note: 内部的には 109 Complex Text
}

# TableDef Attributes 定数 (ビットフラグ)
$dbSystemObject = 0x80000002 # システムオブジェクト (MSys*)
$dbHiddenObject = 0x1        # 隠しオブジェクト
$dbAttachedTable = 0x40000000 # リンクテーブル (ネイティブ Access)
$dbAttachedODBC = 0x20000000 # リンクテーブル (ODBC)
# Relation Attributes 定数
$dbRelationUpdateCascade = 0x100  # 連鎖更新
$dbRelationDeleteCascade = 0x1000 # 連鎖削除
$dbRelationDontEnforce = 0x2      # 参照整合性を設定しない (Enforce = Not ($_.Attributes -band $dbRelationDontEnforce))

if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
    Write-Host " 出力フォルダを作成: $outputFolder"
}

Get-ChildItem -Path $targetFolder -Filter *.accdb | ForEach-Object {
    $accessFile = $_.FullName
    $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($accessFile) + "_Schema.txt"
    $outputFile = Join-Path $outputFolder $outputFileName

    Write-Host " 処理中: $($accessFile)"
    $db = $null # エラー処理用に初期化

    try {
        # データベースを開く (共有モード、読み取り専用)
        $db = $dao.OpenDatabase($accessFile, $false, $true) # ReadOnly = $true
        "=== FILE: $($accessFile) ===" | Out-File $outputFile -Encoding UTF8 
        "Schema generated on: $(Get-Date)`n" | Out-File $outputFile -Encoding UTF8 

        # --- TableDefs を一度取得 ---
        $allTableDefs = $db.TableDefs

        # TABLES (通常のテーブル)
        Write-Host "  - テーブル定義出力中..."
        Add-Content $outputFile -Value "=== TABLE DEFINITIONS ===" -Encoding UTF8
        # 修正: 正しい属性で判定し、システム/隠しオブジェクトを除外
        $tables = $allTableDefs | Where-Object {
            (-not ($_.Attributes -band $dbSystemObject)) -and `
            (-not ($_.Attributes -band $dbHiddenObject)) -and `
            (-not $_.Name.StartsWith("MSys")) -and `
            (-not $_.Name.StartsWith("~")) -and `
            ($_.Attributes -band $dbAttachedTable) -eq 0 -and `
            ($_.Attributes -band $dbAttachedODBC) -eq 0
        } | Sort-Object Name

        Write-Host "    (検出された通常テーブル数: $($tables.Count))"
        if ($tables.Count -eq 0) {
            Add-Content $outputFile -Value "(通常テーブルは見つかりませんでした)" -Encoding UTF8
        }

        foreach ($table in $tables) {
            Add-Content $outputFile -Value "TABLE: $($table.Name)" -Encoding UTF8
            $pkFields = @{} # 主キーフィールドを格納 (複合キー対応のためハッシュテーブル)
            # 主キー情報の取得 (インデックスから)
            try {
                foreach ($index in $table.Indexes) {
                    if ($index.Primary) {
                        # DAO Index Fields はコレクションではない場合があるため、Item() でアクセス
                        for ($i = 0; $i -lt $index.Fields.Count; $i++) {
                            $pkFieldName = $index.Fields.Item($i).Name
                            $pkFields[$pkFieldName] = $true # ハッシュテーブルに追加
                        }
                        break # 主キーは一つのみ
                    }
                }
            }
            catch {
                Write-Warning "    警告: テーブル '$($table.Name)' の主キー情報取得中にエラーが発生しました。 $_"
            }

            # フィールド情報を順序通りに出力
            foreach ($field in $table.Fields | Sort-Object OrdinalPosition) {
                # データ型の取得とマッピング
                $typeName = "Unknown Type#$($field.Type)" # デフォルト値
                if ($typeMap.ContainsKey($field.Type)) {
                    $typeName = $typeMap[$field.Type]
                    # 複数値フィールドの場合、付加情報を取得
                    if ($field.Type -eq 109) {
                        try {
                            $lookupField = $field.Properties("Lookup").Value # これは正しくないかも、複数値は複雑
                            # 実際には Complex Type の判定が必要
                            $typeName += " (Complex)"
                        }
                        catch { $typeName += " (Complex?)" }
                    }
                }

                # サイズ情報
                $sizeInfo = ""
                # Short Text, Binary は Size プロパティが有効
                if (($field.Type -eq 10 -or $field.Type -eq 17) -and $field.Size -gt 0) {
                    $sizeInfo = "($($field.Size))"
                    # Decimal 型は Size (Precision) と DecimalPlaces (Scale)
                }
                elseif ($field.Type -eq 20) {
                    try {
                        $precision = $field.Size # DAOではSizeがPrecision
                        $scale = $field.Properties.Item("DecimalPlaces").Value # DecimalPlacesがScale
                        $sizeInfo = "($precision, $scale)"
                    }
                    catch { $sizeInfo = "(?,?)" } # プロパティ取得失敗時
                }

                # Nullable (Required プロパティ)
                $nullable = if ($field.Required) { "No" } else { "Yes" }
                # Primary Key
                $isPK = if ($pkFields.ContainsKey($field.Name)) { "Yes" } else { "No" }
                # AllowZeroLength (Short Text, Long Text, Binary)
                $allowZeroLength = ""
                if ($field.Type -eq 10 -or $field.Type -eq 12 -or $field.Type -eq 17) {
                    try {
                        $allowZeroLength = if ($field.AllowZeroLength) { ", AllowZeroLength" } else { "" }
                    }
                    catch { $allowZeroLength = "" } # 古いDAO/フィールドタイプでは存在しない場合
                }
                # DefaultValue
                $defaultValue = ""
                try {
                    # DefaultValueが空文字やNullでない場合のみ表示
                    $dv = $field.DefaultValue
                    if ($dv -ne $null -and $dv -ne "") {
                        # 文字列リテラルの場合はクォートを追加 (例: "abc", ="abc")
                        if ($dv -is [string] -and $dv -notlike "'*'" -and $dv -notlike '"*"' -and $dv -notlike "=*") {
                            $defaultValue = ", Default: `"$dv`""
                        }
                        else {
                            $defaultValue = ", Default: $dv"
                        }
                    }
                }
                catch { $defaultValue = "" } # 存在しない場合など

                # 書式 (Format プロパティ)
                $format = ""
                try {
                    $fmt = $field.Properties.Item("Format").Value
                    if ($fmt -ne $null -and $fmt -ne "") {
                        $format = ", Format: $fmt"
                    }
                }
                catch { $format = "" }

                # 定型入力 (InputMask プロパティ)
                $inputMask = ""
                try {
                    $im = $field.Properties.Item("InputMask").Value
                    if ($im -ne $null -and $im -ne "") {
                        $inputMask = ", InputMask: $im"
                    }
                }
                catch { $inputMask = "" }


                $line = "  Field: {0,-25} Type: {1,-20} Nullable: {2,-3} PK: {3,-3}{4}{5}{6}{7}" -f $field.Name, "$typeName$sizeInfo", $nullable, $isPK, $allowZeroLength, $defaultValue, $format, $inputMask
                Add-Content $outputFile -Value $line -Encoding UTF8
            }
            Add-Content $outputFile -Value "" -Encoding UTF8 # テーブルごとに空行
        }

        # LINKED TABLES
        Write-Host "  - リンクテーブル出力中..."
        Add-Content $outputFile -Value "=== LINKED TABLES ===" -Encoding UTF8
        # 修正: 正しい属性で判定し、システム/隠しオブジェクトを除外
        $linkedTables = $allTableDefs | Where-Object {
            (-not ($_.Attributes -band $dbSystemObject)) -and `
            (-not ($_.Attributes -band $dbHiddenObject)) -and `
            (-not $_.Name.StartsWith("MSys")) -and `
            (-not $_.Name.StartsWith("~")) -and `
            (($_.Attributes -band $dbAttachedTable) -ne 0 -or ($_.Attributes -band $dbAttachedODBC) -ne 0)
        } | Sort-Object Name

        Write-Host "    (検出されたリンクテーブル数: $($linkedTables.Count))"
        if ($linkedTables.Count -eq 0) {
            Add-Content $outputFile -Value "(リンクテーブルは見つかりませんでした)" -Encoding UTF8
        }

        foreach ($table in $linkedTables) {
            Add-Content $outputFile -Value "LinkedTable: $($table.Name)" -Encoding UTF8
            $linkType = if (($table.Attributes -band $dbAttachedODBC) -ne 0) { "ODBC" } elseif (($table.Attributes -band $dbAttachedTable) -ne 0) { "Access" } else { "Unknown" }
            Add-Content $outputFile -Value "  Type: $linkType" -Encoding UTF8
            Add-Content $outputFile -Value "  Connect: $($table.Connect)" -Encoding UTF8
            Add-Content $outputFile -Value "  SourceTable: $($table.SourceTableName)" -Encoding UTF8

            # リンクテーブルのフィールド情報は、リンクが有効でないと取得できない
            try {
                # 試しにフィールド数を取得 (失敗したらcatchへ)
                $fieldCount = $table.Fields.Count
                if ($fieldCount -gt 0) {
                    Add-Content $outputFile -Value "  Fields: (Link OK)" -Encoding UTF8
                    foreach ($field in $table.Fields | Sort-Object OrdinalPosition) {
                        $typeName = "Unknown Type#$($field.Type)"
                        if ($typeMap.ContainsKey($field.Type)) { $typeName = $typeMap[$field.Type] }

                        $sizeInfo = ""
                        if (($field.Type -eq 10 -or $field.Type -eq 17) -and $field.Size -gt 0) { $sizeInfo = "($($field.Size))" }
                        elseif ($field.Type -eq 20) {
                            try { $precision = $field.Size; $scale = $field.Properties.Item("DecimalPlaces").Value; $sizeInfo = "($precision, $scale)" }
                            catch { $sizeInfo = "(?,?)" }
                        }

                        $nullable = if ($field.Required) { "No" } else { "Yes" }
                        $line = "    Field: {0,-25} Type: {1,-20} Nullable: {2}" -f $field.Name, "$typeName$sizeInfo", $nullable
                        Add-Content $outputFile -Value $line -Encoding UTF8
                    }
                }
                else {
                    # フィールド数0の場合は、リンクはOKだが空の可能性
                    Add-Content $outputFile -Value "  Fields: (Link OK, but no fields found or empty source)" -Encoding UTF8
                }
            }
            catch {
                Add-Content $outputFile -Value "  Fields: ※ Link Broken or Permissions Issue - Cannot retrieve field info." -Encoding UTF8
                Write-Warning "    警告: リンクテーブル '$($table.Name)' のフィールド情報取得に失敗しました (リンク切れ/権限不足の可能性)。詳細: $($_.Exception.Message)"
            }
            Add-Content $outputFile -Value "" -Encoding UTF8 # リンクテーブルごとに空行
        }

        # RELATIONS
        Write-Host "  - リレーション出力中..."
        Add-Content $outputFile -Value "=== RELATIONSHIPS ===" -Encoding UTF8
        $relations = $db.Relations | Sort-Object Name
        Write-Host "    (検出されたリレーションシップ数: $($relations.Count))"
        if ($relations.Count -eq 0) {
            Add-Content $outputFile -Value "(リレーションシップは見つかりませんでした)" -Encoding UTF8
        }
        foreach ($rel in $relations) {
            if ($rel -ne $null -and $rel.Name -ne $null) {
                Add-Content $outputFile -Value "Relation: $($rel.Name)" -Encoding UTF8
                if ($rel.Fields -ne $null -and $rel.Fields.Count -gt 0) {
                    # 複合キーに対応してフィールドを列挙
                    $localFields = ($rel.Fields | ForEach-Object { $_.Name }) -join ", "
                    $foreignFields = ($rel.Fields | ForEach-Object { $_.ForeignName }) -join ", "
                    Add-Content $outputFile -Value "  From: $($rel.Table) ($localFields)" -Encoding UTF8
                    Add-Content $outputFile -Value "  To:   $($rel.ForeignTable) ($foreignFields)" -Encoding UTF8
                }
                else {
                    Add-Content $outputFile -Value "  From: $($rel.Table) (Fields info missing)" -Encoding UTF8
                    Add-Content $outputFile -Value "  To:   $($rel.ForeignTable) (Fields info missing)" -Encoding UTF8
                }
                # 参照整合性とオプション
                $isEnforced = (-not ($rel.Attributes -band $dbRelationDontEnforce))
                $cascadeUpdate = ($rel.Attributes -band $dbRelationUpdateCascade) -ne 0
                $cascadeDelete = ($rel.Attributes -band $dbRelationDeleteCascade) -ne 0
                # Join Type (0: Inner, 1: Left, 2: Right) - これはクエリのプロパティだが、リレーションにもある場合がある
                $joinTypeVal = try { $rel.Attributes -band 0x03 } catch { 0 } # 通常は 0
                $joinType = switch ($joinTypeVal) {
                    1 { "LEFT" }
                    2 { "RIGHT" }
                    default { "INNER" }
                }
                Add-Content $outputFile -Value "  Integrity: Enforced=$isEnforced, CascadeUpdate=$cascadeUpdate, CascadeDelete=$cascadeDelete, JoinType=$joinType" -Encoding UTF8
                Add-Content $outputFile -Value "" -Encoding UTF8
            }
            else {
                Write-Warning "    警告: 無効なリレーションシップオブジェクトが見つかりました。"
            }
        }

        # QUERIES
        Write-Host "  - クエリ出力中..."
        Add-Content $outputFile -Value "=== QUERIES ===" -Encoding UTF8
        # 一時クエリ (~ で始まる) を除外
        $queries = $db.QueryDefs | Where-Object { -not $_.Name.StartsWith("~") } | Sort-Object Name
        Write-Host "    (検出されたクエリ数: $($queries.Count))"
        if ($queries.Count -eq 0) {
            Add-Content $outputFile -Value "(クエリは見つかりませんでした)" -Encoding UTF8
        }
        foreach ($query in $queries) {
            Add-Content $outputFile -Value "Query: $($query.Name)" -Encoding UTF8
            # クエリタイプを取得
            $queryType = "Unknown ($($query.Type))"
            switch ($query.Type) {
                0 { $queryType = "Select" }
                16 { $queryType = "Crosstab" }
                32 { $queryType = "Delete Action" }
                48 { $queryType = "Update Action" }
                64 { $queryType = "Append Action" }
                80 { $queryType = "Make-Table Action" }
                96 { $queryType = "Data Definition (DDL)" }
                112 { $queryType = "Pass-Through" }
                128 { $queryType = "Union" }
                # その他: 240=ptAppend, 208=ptMakeTable などODBCパススルー系アクション
            }
            Add-Content $outputFile -Value ("  Type: " + $queryType) -Encoding UTF8
            # ReturnsRecords プロパティ (Select系かAction系か)
            try {
                Add-Content $outputFile -Value ("  ReturnsRecords: " + $query.ReturnsRecords) -Encoding UTF8
            }
            catch {}
            # SQL を取得 (改行を維持しつつ整形)
            $sql = $query.SQL # 基本的にそのまま取得
            # 必要であれば $sql = $query.SQL -replace '[\r\n]+', ' ' -replace '\s+', ' '
            Add-Content $outputFile -Value ("  SQL: " + $sql) -Encoding UTF8

            # パラメータ情報
            try {
                if ($query.Parameters.Count -gt 0) {
                    Add-Content $outputFile -Value "  Parameters:" -Encoding UTF8
                    foreach ($param in $query.Parameters) {
                        $paramTypeName = "Unknown Type#$($param.Type)"
                        if ($typeMap.ContainsKey($param.Type)) { $paramTypeName = $typeMap[$param.Type] }
                        Add-Content $outputFile -Value ("    Param: {0,-20} Type: {1}" -f $param.Name, $paramTypeName) -Encoding UTF8
                    }
                }
            }
            catch {
                # パラメータがないクエリでは Parameters プロパティアクセスでエラーになる場合がある
            }

            Add-Content $outputFile -Value "" -Encoding UTF8 # クエリごとに空行
        }

        Write-Host " 完了: $($outputFile)"

    }
    catch {
        Write-Warning " エラー処理中: $accessFile"
        Write-Warning "  エラー詳細: $($_.Exception.Message)"
        Write-Warning "  スタックトレース: $($_.ScriptStackTrace)"
        # エラー内容を出力ファイルにも追記
        "--- ERROR ---" | Out-File $outputFile -Encoding UTF8 
        "Error processing file: $accessFile" | Out-File $outputFile -Encoding UTF8 
        "$($_.Exception.Message)" | Out-File $outputFile -Encoding UTF8 
        "$($_.ScriptStackTrace)" | Out-File $outputFile -Encoding UTF8 
        "---------------" | Out-File $outputFile -Encoding UTF8 
    }
    finally {
        # DBオブジェクトが作成されていれば閉じて解放
        if ($db -ne $null) {
            try { $db.Close() } catch { Write-Warning "DB Close中にエラー: $($_.Exception.Message)" }
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($db)
            $db = $null
        }
        # ガベージコレクションを強制 (COMオブジェクト解放のため)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# DAOエンジンの解放
if ($dao -ne $null) {
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dao)
    $dao = $null
}
# 再度ガベージコレクション
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "全ての処理が完了しました！"