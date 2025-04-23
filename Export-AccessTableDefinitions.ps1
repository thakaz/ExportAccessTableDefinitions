# <#
# .SYNOPSIS
# Access �f�[�^�x�[�X (.accdb) �̃X�L�[�}�����e�L�X�g�t�@�C���ɏo�͂��܂��B
# .DESCRIPTION
# �w�肳�ꂽ�t�H���_���̂��ׂĂ� .accdb �t�@�C�����������A���ꂼ��̃f�[�^�x�[�X�ɂ��āA
# �e�[�u����`�A�����N�e�[�u�����A�����[�V�����V�b�v�A�N�G����`�𒊏o���A
# "SchemaOutput" �t�H���_���Ƀt�@�C�������Ƃ̃X�L�[�}�t�@�C�� (.txt) ���쐬���܂��B
#
# �ˑ��֌W: Microsoft Data Access Objects (DAO) ���K�v�ł��B�ʏ�AMicrosoft Office/Access
#          ���C���X�g�[������Ă���Η��p�\�ł��B
# #>
Write-Host "Access �X�L�[�}�o�̓X�N���v�g���J�n���܂�..."

$targetFolder = $PSScriptRoot
$outputFolder = Join-Path $PSScriptRoot "SchemaOutput"

# DAO �I�u�W�F�N�g���쐬 (Access 2007 �ȍ~�� .accdb �`���ɑΉ�)
try {
    $dao = New-Object -ComObject DAO.DBEngine.120
    Write-Host "INFO: DAO DBEngine Version $($dao.Version) ���g�p���܂��B"
}
catch {
    Write-Error "DAO �I�u�W�F�N�g (DAO.DBEngine.120) �̍쐬�Ɏ��s���܂����BMicrosoft Access �܂��� Access Database Engine ���C���X�g�[������Ă��邩�m�F���Ă��������B"
    exit 1
}


# Access �f�[�^�^�ƈ�ʓI�Ȗ��̂̃}�b�s���O (���ڍׂ�)
$typeMap = @{
    1   = "Boolean"; # Yes/No
    2   = "Byte"; # ���l�^ (�o�C�g�^)
    3   = "Integer"; # ���l�^ (�����^)
    4   = "Long"; # ���l�^ (�������^)
    5   = "Currency"; # �ʉ݌^
    6   = "Single"; # ���l�^ (�P���x���������_���^)
    7   = "Double"; # ���l�^ (�{���x���������_���^)
    8   = "Date/Time"; # ���t/�����^
    10  = "Short Text"; # �Z���e�L�X�g (�ȑO�� Text)
    11  = "OLE Object"; # OLE �I�u�W�F�N�g
    12  = "Long Text"; # �����e�L�X�g (�ȑO�� Memo)
    15  = "Replication ID"; # ���v���P�[�V���� ID (GUID)
    16  = "BigInt"; # �傫�����l (Access 2016+)
    17  = "Binary"; # �o�C�i���f�[�^ (VarBinary �Ȃ�)
    18  = "Date/Time Extended"; # ���t/�����g���^ (Access 2016+)
    20  = "Decimal"; # ���l�^ (�\�i�^) - DAO�ł�Numeric�Ƃ�
    101 = "Attachment"; # �Y�t�t�@�C��
    109 = "Multi-valued"; # �����l�����t�B�[���h (Complex Text�Ȃ�) - Note: �����I�ɂ� 109 Complex Text
}

# TableDef Attributes �萔 (�r�b�g�t���O)
$dbSystemObject = 0x80000002 # �V�X�e���I�u�W�F�N�g (MSys*)
$dbHiddenObject = 0x1        # �B���I�u�W�F�N�g
$dbAttachedTable = 0x40000000 # �����N�e�[�u�� (�l�C�e�B�u Access)
$dbAttachedODBC = 0x20000000 # �����N�e�[�u�� (ODBC)
# Relation Attributes �萔
$dbRelationUpdateCascade = 0x100  # �A���X�V
$dbRelationDeleteCascade = 0x1000 # �A���폜
$dbRelationDontEnforce = 0x2      # �Q�Ɛ�������ݒ肵�Ȃ� (Enforce = Not ($_.Attributes -band $dbRelationDontEnforce))

if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
    Write-Host " �o�̓t�H���_���쐬: $outputFolder"
}

Get-ChildItem -Path $targetFolder -Filter *.accdb | ForEach-Object {
    $accessFile = $_.FullName
    $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($accessFile) + "_Schema.txt"
    $outputFile = Join-Path $outputFolder $outputFileName

    Write-Host " ������: $($accessFile)"
    $db = $null # �G���[�����p�ɏ�����

    try {
        # �f�[�^�x�[�X���J�� (���L���[�h�A�ǂݎ���p)
        $db = $dao.OpenDatabase($accessFile, $false, $true) # ReadOnly = $true
        "=== FILE: $($accessFile) ===" | Out-File $outputFile -Encoding UTF8 
        "Schema generated on: $(Get-Date)`n" | Out-File $outputFile -Encoding UTF8 

        # --- TableDefs ����x�擾 ---
        $allTableDefs = $db.TableDefs

        # TABLES (�ʏ�̃e�[�u��)
        Write-Host "  - �e�[�u����`�o�͒�..."
        Add-Content $outputFile -Value "=== TABLE DEFINITIONS ===" -Encoding UTF8
        # �C��: �����������Ŕ��肵�A�V�X�e��/�B���I�u�W�F�N�g�����O
        $tables = $allTableDefs | Where-Object {
            (-not ($_.Attributes -band $dbSystemObject)) -and `
            (-not ($_.Attributes -band $dbHiddenObject)) -and `
            (-not $_.Name.StartsWith("MSys")) -and `
            (-not $_.Name.StartsWith("~")) -and `
            ($_.Attributes -band $dbAttachedTable) -eq 0 -and `
            ($_.Attributes -band $dbAttachedODBC) -eq 0
        } | Sort-Object Name

        Write-Host "    (���o���ꂽ�ʏ�e�[�u����: $($tables.Count))"
        if ($tables.Count -eq 0) {
            Add-Content $outputFile -Value "(�ʏ�e�[�u���͌�����܂���ł���)" -Encoding UTF8
        }

        foreach ($table in $tables) {
            Add-Content $outputFile -Value "TABLE: $($table.Name)" -Encoding UTF8
            $pkFields = @{} # ��L�[�t�B�[���h���i�[ (�����L�[�Ή��̂��߃n�b�V���e�[�u��)
            # ��L�[���̎擾 (�C���f�b�N�X����)
            try {
                foreach ($index in $table.Indexes) {
                    if ($index.Primary) {
                        # DAO Index Fields �̓R���N�V�����ł͂Ȃ��ꍇ�����邽�߁AItem() �ŃA�N�Z�X
                        for ($i = 0; $i -lt $index.Fields.Count; $i++) {
                            $pkFieldName = $index.Fields.Item($i).Name
                            $pkFields[$pkFieldName] = $true # �n�b�V���e�[�u���ɒǉ�
                        }
                        break # ��L�[�͈�̂�
                    }
                }
            }
            catch {
                Write-Warning "    �x��: �e�[�u�� '$($table.Name)' �̎�L�[���擾���ɃG���[���������܂����B $_"
            }

            # �t�B�[���h���������ʂ�ɏo��
            foreach ($field in $table.Fields | Sort-Object OrdinalPosition) {
                # �f�[�^�^�̎擾�ƃ}�b�s���O
                $typeName = "Unknown Type#$($field.Type)" # �f�t�H���g�l
                if ($typeMap.ContainsKey($field.Type)) {
                    $typeName = $typeMap[$field.Type]
                    # �����l�t�B�[���h�̏ꍇ�A�t�������擾
                    if ($field.Type -eq 109) {
                        try {
                            $lookupField = $field.Properties("Lookup").Value # ����͐������Ȃ������A�����l�͕��G
                            # ���ۂɂ� Complex Type �̔��肪�K�v
                            $typeName += " (Complex)"
                        }
                        catch { $typeName += " (Complex?)" }
                    }
                }

                # �T�C�Y���
                $sizeInfo = ""
                # Short Text, Binary �� Size �v���p�e�B���L��
                if (($field.Type -eq 10 -or $field.Type -eq 17) -and $field.Size -gt 0) {
                    $sizeInfo = "($($field.Size))"
                    # Decimal �^�� Size (Precision) �� DecimalPlaces (Scale)
                }
                elseif ($field.Type -eq 20) {
                    try {
                        $precision = $field.Size # DAO�ł�Size��Precision
                        $scale = $field.Properties.Item("DecimalPlaces").Value # DecimalPlaces��Scale
                        $sizeInfo = "($precision, $scale)"
                    }
                    catch { $sizeInfo = "(?,?)" } # �v���p�e�B�擾���s��
                }

                # Nullable (Required �v���p�e�B)
                $nullable = if ($field.Required) { "No" } else { "Yes" }
                # Primary Key
                $isPK = if ($pkFields.ContainsKey($field.Name)) { "Yes" } else { "No" }
                # AllowZeroLength (Short Text, Long Text, Binary)
                $allowZeroLength = ""
                if ($field.Type -eq 10 -or $field.Type -eq 12 -or $field.Type -eq 17) {
                    try {
                        $allowZeroLength = if ($field.AllowZeroLength) { ", AllowZeroLength" } else { "" }
                    }
                    catch { $allowZeroLength = "" } # �Â�DAO/�t�B�[���h�^�C�v�ł͑��݂��Ȃ��ꍇ
                }
                # DefaultValue
                $defaultValue = ""
                try {
                    # DefaultValue���󕶎���Null�łȂ��ꍇ�̂ݕ\��
                    $dv = $field.DefaultValue
                    if ($dv -ne $null -and $dv -ne "") {
                        # �����񃊃e�����̏ꍇ�̓N�H�[�g��ǉ� (��: "abc", ="abc")
                        if ($dv -is [string] -and $dv -notlike "'*'" -and $dv -notlike '"*"' -and $dv -notlike "=*") {
                            $defaultValue = ", Default: `"$dv`""
                        }
                        else {
                            $defaultValue = ", Default: $dv"
                        }
                    }
                }
                catch { $defaultValue = "" } # ���݂��Ȃ��ꍇ�Ȃ�

                # ���� (Format �v���p�e�B)
                $format = ""
                try {
                    $fmt = $field.Properties.Item("Format").Value
                    if ($fmt -ne $null -and $fmt -ne "") {
                        $format = ", Format: $fmt"
                    }
                }
                catch { $format = "" }

                # ��^���� (InputMask �v���p�e�B)
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
            Add-Content $outputFile -Value "" -Encoding UTF8 # �e�[�u�����Ƃɋ�s
        }

        # LINKED TABLES
        Write-Host "  - �����N�e�[�u���o�͒�..."
        Add-Content $outputFile -Value "=== LINKED TABLES ===" -Encoding UTF8
        # �C��: �����������Ŕ��肵�A�V�X�e��/�B���I�u�W�F�N�g�����O
        $linkedTables = $allTableDefs | Where-Object {
            (-not ($_.Attributes -band $dbSystemObject)) -and `
            (-not ($_.Attributes -band $dbHiddenObject)) -and `
            (-not $_.Name.StartsWith("MSys")) -and `
            (-not $_.Name.StartsWith("~")) -and `
            (($_.Attributes -band $dbAttachedTable) -ne 0 -or ($_.Attributes -band $dbAttachedODBC) -ne 0)
        } | Sort-Object Name

        Write-Host "    (���o���ꂽ�����N�e�[�u����: $($linkedTables.Count))"
        if ($linkedTables.Count -eq 0) {
            Add-Content $outputFile -Value "(�����N�e�[�u���͌�����܂���ł���)" -Encoding UTF8
        }

        foreach ($table in $linkedTables) {
            Add-Content $outputFile -Value "LinkedTable: $($table.Name)" -Encoding UTF8
            $linkType = if (($table.Attributes -band $dbAttachedODBC) -ne 0) { "ODBC" } elseif (($table.Attributes -band $dbAttachedTable) -ne 0) { "Access" } else { "Unknown" }
            Add-Content $outputFile -Value "  Type: $linkType" -Encoding UTF8
            Add-Content $outputFile -Value "  Connect: $($table.Connect)" -Encoding UTF8
            Add-Content $outputFile -Value "  SourceTable: $($table.SourceTableName)" -Encoding UTF8

            # �����N�e�[�u���̃t�B�[���h���́A�����N���L���łȂ��Ǝ擾�ł��Ȃ�
            try {
                # �����Ƀt�B�[���h�����擾 (���s������catch��)
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
                    # �t�B�[���h��0�̏ꍇ�́A�����N��OK������̉\��
                    Add-Content $outputFile -Value "  Fields: (Link OK, but no fields found or empty source)" -Encoding UTF8
                }
            }
            catch {
                Add-Content $outputFile -Value "  Fields: �� Link Broken or Permissions Issue - Cannot retrieve field info." -Encoding UTF8
                Write-Warning "    �x��: �����N�e�[�u�� '$($table.Name)' �̃t�B�[���h���擾�Ɏ��s���܂��� (�����N�؂�/�����s���̉\��)�B�ڍ�: $($_.Exception.Message)"
            }
            Add-Content $outputFile -Value "" -Encoding UTF8 # �����N�e�[�u�����Ƃɋ�s
        }

        # RELATIONS
        Write-Host "  - �����[�V�����o�͒�..."
        Add-Content $outputFile -Value "=== RELATIONSHIPS ===" -Encoding UTF8
        $relations = $db.Relations | Sort-Object Name
        Write-Host "    (���o���ꂽ�����[�V�����V�b�v��: $($relations.Count))"
        if ($relations.Count -eq 0) {
            Add-Content $outputFile -Value "(�����[�V�����V�b�v�͌�����܂���ł���)" -Encoding UTF8
        }
        foreach ($rel in $relations) {
            if ($rel -ne $null -and $rel.Name -ne $null) {
                Add-Content $outputFile -Value "Relation: $($rel.Name)" -Encoding UTF8
                if ($rel.Fields -ne $null -and $rel.Fields.Count -gt 0) {
                    # �����L�[�ɑΉ����ăt�B�[���h���
                    $localFields = ($rel.Fields | ForEach-Object { $_.Name }) -join ", "
                    $foreignFields = ($rel.Fields | ForEach-Object { $_.ForeignName }) -join ", "
                    Add-Content $outputFile -Value "  From: $($rel.Table) ($localFields)" -Encoding UTF8
                    Add-Content $outputFile -Value "  To:   $($rel.ForeignTable) ($foreignFields)" -Encoding UTF8
                }
                else {
                    Add-Content $outputFile -Value "  From: $($rel.Table) (Fields info missing)" -Encoding UTF8
                    Add-Content $outputFile -Value "  To:   $($rel.ForeignTable) (Fields info missing)" -Encoding UTF8
                }
                # �Q�Ɛ������ƃI�v�V����
                $isEnforced = (-not ($rel.Attributes -band $dbRelationDontEnforce))
                $cascadeUpdate = ($rel.Attributes -band $dbRelationUpdateCascade) -ne 0
                $cascadeDelete = ($rel.Attributes -band $dbRelationDeleteCascade) -ne 0
                # Join Type (0: Inner, 1: Left, 2: Right) - ����̓N�G���̃v���p�e�B�����A�����[�V�����ɂ�����ꍇ������
                $joinTypeVal = try { $rel.Attributes -band 0x03 } catch { 0 } # �ʏ�� 0
                $joinType = switch ($joinTypeVal) {
                    1 { "LEFT" }
                    2 { "RIGHT" }
                    default { "INNER" }
                }
                Add-Content $outputFile -Value "  Integrity: Enforced=$isEnforced, CascadeUpdate=$cascadeUpdate, CascadeDelete=$cascadeDelete, JoinType=$joinType" -Encoding UTF8
                Add-Content $outputFile -Value "" -Encoding UTF8
            }
            else {
                Write-Warning "    �x��: �����ȃ����[�V�����V�b�v�I�u�W�F�N�g��������܂����B"
            }
        }

        # QUERIES
        Write-Host "  - �N�G���o�͒�..."
        Add-Content $outputFile -Value "=== QUERIES ===" -Encoding UTF8
        # �ꎞ�N�G�� (~ �Ŏn�܂�) �����O
        $queries = $db.QueryDefs | Where-Object { -not $_.Name.StartsWith("~") } | Sort-Object Name
        Write-Host "    (���o���ꂽ�N�G����: $($queries.Count))"
        if ($queries.Count -eq 0) {
            Add-Content $outputFile -Value "(�N�G���͌�����܂���ł���)" -Encoding UTF8
        }
        foreach ($query in $queries) {
            Add-Content $outputFile -Value "Query: $($query.Name)" -Encoding UTF8
            # �N�G���^�C�v���擾
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
                # ���̑�: 240=ptAppend, 208=ptMakeTable �Ȃ�ODBC�p�X�X���[�n�A�N�V����
            }
            Add-Content $outputFile -Value ("  Type: " + $queryType) -Encoding UTF8
            # ReturnsRecords �v���p�e�B (Select�n��Action�n��)
            try {
                Add-Content $outputFile -Value ("  ReturnsRecords: " + $query.ReturnsRecords) -Encoding UTF8
            }
            catch {}
            # SQL ���擾 (���s���ێ������`)
            $sql = $query.SQL # ��{�I�ɂ��̂܂܎擾
            # �K�v�ł���� $sql = $query.SQL -replace '[\r\n]+', ' ' -replace '\s+', ' '
            Add-Content $outputFile -Value ("  SQL: " + $sql) -Encoding UTF8

            # �p�����[�^���
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
                # �p�����[�^���Ȃ��N�G���ł� Parameters �v���p�e�B�A�N�Z�X�ŃG���[�ɂȂ�ꍇ������
            }

            Add-Content $outputFile -Value "" -Encoding UTF8 # �N�G�����Ƃɋ�s
        }

        Write-Host " ����: $($outputFile)"

    }
    catch {
        Write-Warning " �G���[������: $accessFile"
        Write-Warning "  �G���[�ڍ�: $($_.Exception.Message)"
        Write-Warning "  �X�^�b�N�g���[�X: $($_.ScriptStackTrace)"
        # �G���[���e���o�̓t�@�C���ɂ��ǋL
        "--- ERROR ---" | Out-File $outputFile -Encoding UTF8 
        "Error processing file: $accessFile" | Out-File $outputFile -Encoding UTF8 
        "$($_.Exception.Message)" | Out-File $outputFile -Encoding UTF8 
        "$($_.ScriptStackTrace)" | Out-File $outputFile -Encoding UTF8 
        "---------------" | Out-File $outputFile -Encoding UTF8 
    }
    finally {
        # DB�I�u�W�F�N�g���쐬����Ă���Ε��ĉ��
        if ($db -ne $null) {
            try { $db.Close() } catch { Write-Warning "DB Close���ɃG���[: $($_.Exception.Message)" }
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($db)
            $db = $null
        }
        # �K�x�[�W�R���N�V���������� (COM�I�u�W�F�N�g����̂���)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# DAO�G���W���̉��
if ($dao -ne $null) {
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dao)
    $dao = $null
}
# �ēx�K�x�[�W�R���N�V����
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "�S�Ă̏������������܂����I"