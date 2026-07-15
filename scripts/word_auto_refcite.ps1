<#
.SYNOPSIS
将正文顺序编码引文写为 Word 原生“编号项”交叉引用。

.DESCRIPTION
仅适用于 Windows 上已安装 Microsoft Word 的环境。文后列表必须已是 Word 自动编号；
本脚本调用 Word 的 InsertCrossReference，而不手写 REF 域，确保 Ctrl+单击跳转和域更新兼容。
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$referenceTitles = @('参考文献', '〖参考文献〗', '主要参考文献', 'References', 'Bibliography')
$citationPattern = '\[(\d+(?:\s*[,，、\-–—]\s*\d+)*)\]'

function Get-ReferenceHeadingStart {
    param($Document)
    foreach ($paragraph in $Document.Paragraphs) {
        $text = $paragraph.Range.Text.Trim([char]13, [char]7)
        if ($referenceTitles -contains $text) {
            return $paragraph.Range.Start
        }
    }
    throw '未找到参考文献标题。'
}

function Get-InsertedReferenceField {
    param($Document, [int]$SearchStart)
    $candidate = $null
    for ($index = 1; $index -le $Document.Fields.Count; $index++) {
        $field = $Document.Fields.Item($index)
        if ($field.Code.Text -notmatch 'REF Ref_') {
            continue
        }
        if ($field.Result.Start -lt $SearchStart) {
            continue
        }
        if ($null -eq $candidate -or $field.Result.Start -lt $candidate.Result.Start) {
            $candidate = $field
        }
    }
    if ($null -eq $candidate) {
        throw 'Word 未返回新建的编号项交叉引用字段。'
    }
    return $candidate
}

function Set-CitationFormat {
    param($Range)
    $Range.Font.Superscript = 1
    $Range.Font.Size = 9
    $Range.Font.Color = 0
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    throw "输入文件不存在：$InputPath"
}
if (-not $OutputPath) {
    $OutputPath = Join-Path (Split-Path -LiteralPath $InputPath -Parent) (([IO.Path]::GetFileNameWithoutExtension($InputPath)) + '_自动编号交叉引用.docx')
}
if ((Resolve-Path -LiteralPath $InputPath).Path -eq [IO.Path]::GetFullPath($OutputPath)) {
    throw '输出文件不能覆盖输入文件，请指定新的 --Output。'
}

Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($OutputPath, $false, $false)

    $referenceStart = Get-ReferenceHeadingStart $document
    $body = $document.Range(0, $referenceStart)
    $bodyText = $body.Text
    $referenceItems = $document.GetCrossReferenceItems(0)
    if ($referenceItems.Count -eq 0) {
        throw '参考文献区未识别到 Word 自动编号条目。'
    }
    $matches = [regex]::Matches($bodyText, $citationPattern)
    if ($matches.Count -eq 0) {
        throw '正文未找到 [N]、[N,N] 或 [N-N] 引文。'
    }
    # Word 的 Range 位置与 .NET 字符串位置在少数 Unicode 字符下不一致，因此用 Word 查找记录真实范围。
    $citationRanges = @()
    $searchStart = $body.Start
    foreach ($match in $matches) {
        $citation = $document.Range($searchStart, $referenceStart)
        $citation.Find.ClearFormatting()
        $citation.Find.Text = $match.Value
        $citation.Find.Forward = $true
        $citation.Find.Wrap = 0
        $citation.Find.MatchWildcards = $false
        if (-not $citation.Find.Execute()) {
            throw "Word 未能定位正文引文：$($match.Value)"
        }
        $citationRanges += [PSCustomObject]@{
            Start = $citation.Start
            End = $citation.End
            Text = $match.Value
        }
        $searchStart = $citation.End
    }

    # 先建立可见书签，供用户在 Word“交叉引用 -> 书签”窗口中继续维护。
    if ($document.ListParagraphs.Count -ne $referenceItems.Count) {
        throw '文档中存在非参考文献的自动列表，无法安全映射编号项。'
    }
    for ($number = 1; $number -le $referenceItems.Count; $number++) {
        $name = 'Ref_{0:D3}' -f $number
        if ($document.Bookmarks.Exists($name)) {
            continue
        }
        $target = $document.ListParagraphs.Item($number).Range.Duplicate
        $target.End = $target.End - 1
        [void]$document.Bookmarks.Add($name, $target)
    }

    $fieldCount = 0
    # 从后向前替换，避免新字段的隐藏代码改变尚未处理的文本位置。
    for ($matchIndex = $citationRanges.Count - 1; $matchIndex -ge 0; $matchIndex--) {
        $citation = $citationRanges[$matchIndex]
        $clusterStart = $citation.Start
        $clusterEnd = $citation.End
        $cluster = $document.Range($clusterStart, $clusterEnd)
        $clusterText = $citation.Text
        $numbers = [regex]::Matches($clusterText, '\d+')
        foreach ($numberMatch in $numbers) {
            $number = [int]$numberMatch.Value
            if ($number -lt 1 -or $number -gt $referenceItems.Count) {
                throw "正文引用了不存在的自动编号条目：[$number]"
            }
        }

        # Word 自动编号交叉引用自身含有方括号，故整体移除原引文后逐项插入。
        $cluster.Text = ''
        $cursor = $clusterStart
        for ($numberIndex = 0; $numberIndex -lt $numbers.Count; $numberIndex++) {
            $number = [int]$numbers.Item($numberIndex).Value
            $insertion = $document.Range($cursor, $cursor)
            $insertion.InsertCrossReference(0, -3, $number, $true, $false, $false, '')
            $field = Get-InsertedReferenceField $document $cursor
            Set-CitationFormat $field.Result
            $fieldCount++

            if ($numberIndex -lt $numbers.Count - 1) {
                $separatorStart = $numbers.Item($numberIndex).Index + $numbers.Item($numberIndex).Length
                $separatorEnd = $numbers.Item($numberIndex + 1).Index
                $separator = $clusterText.Substring($separatorStart, $separatorEnd - $separatorStart)
                $separatorRange = $document.Range($field.Result.End, $field.Result.End)
                $separatorRange.Text = $separator
                Set-CitationFormat $separatorRange
                $cursor = $field.Result.End + $separator.Length
            }
        }
    }

    [void]$document.Fields.Update()
    $actualFields = @()
    for ($index = 1; $index -le $document.Fields.Count; $index++) {
        $field = $document.Fields.Item($index)
        if ($field.Code.Text -match 'REF Ref_') {
            Set-CitationFormat $field.Result
            $actualFields += $field
        }
    }
    if ($actualFields.Count -ne $fieldCount) {
        throw "编号项交叉引用数量异常：预期 $fieldCount，实际 $($actualFields.Count)。"
    }
    if (@($actualFields | Where-Object { $_.Result.Text -like '错误!*' }).Count) {
        throw 'Word 更新后存在无法解析的交叉引用。'
    }
    if (@($actualFields | Where-Object {
        $_.Result.Font.Superscript -eq 0 -or $_.Result.Font.Size -ne 9 -or $_.Result.Font.Color -ne 0
    }).Count) {
        throw 'Word 编号项交叉引用未统一为黑色 9 磅上标。'
    }
    $document.Save()
    Write-Output "已写入 $fieldCount 个 Word 原生编号项交叉引用：$OutputPath"
}
finally {
    if ($null -ne $document) {
        $document.Close($false)
    }
    if ($null -ne $word) {
        $word.Quit()
    }
}







