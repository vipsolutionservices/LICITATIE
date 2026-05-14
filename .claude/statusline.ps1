param()
# Statusline pentru Claude Code:
#   [<alias>] <icon> <subject curent in_progress sau ultim completat>
# Primește pe stdin JSON cu { session_id, transcript_path, cwd, model, version, ... }
# Returnează pe stdout o singură linie text.

$ErrorActionPreference = 'SilentlyContinue'

# Forțăm UTF-8 pentru input și output (altfel ✓/▶ ies deformate pe Windows).
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Citește JSON-ul de pe stdin
$raw = [Console]::In.ReadToEnd()
$alias = '??'
$line  = "[$alias] (no tasks)"

try {
    $ctx = $raw | ConvertFrom-Json
    $tp  = $ctx.transcript_path
    if (-not (Test-Path $tp)) { Write-Output $line; exit 0 }

    # Citește alias-ul din mapping session_id -> alias (per-sesiune)
    $aliasMap = Join-Path $PSScriptRoot '_aliases.json'
    if (Test-Path $aliasMap) {
        try {
            $map = Get-Content $aliasMap -Raw | ConvertFrom-Json
            $sid = "$($ctx.session_id)"
            if ($sid -and $map.PSObject.Properties[$sid]) {
                $alias = $map.PSObject.Properties[$sid].Value
            }
        } catch {}
    }
    $line = "[$alias] (no tasks)"

    # Parcurg jsonl și păstrez ultima stare cunoscută pentru fiecare taskId
    # Evenimente relevante: tool_use cu name=TaskCreate (id, subject) și name=TaskUpdate (taskId, status, subject?)
    # subject definit la create + posibil override la update; status pornește pending la create.
    $tasks = @{}            # id -> @{ subject; status; updatedAt }
    $createOrder = @{}       # id -> int (ordinea apariției TaskCreate, pt fallback nr task)
    $createIdx = 0
    # Pentru TaskCreate, id-ul e returnat în tool_result (mesaj din rolul user ulterior), deci urmărim
    # toolUseId-ul și apoi mapăm rezultatul.
    $pendingCreate = @{}     # tool_use_id -> subject

    Get-Content $tp -ErrorAction Stop | ForEach-Object {
        $evt = $_ | ConvertFrom-Json -ErrorAction SilentlyContinue
        if (-not $evt) { return }
        $msg = $evt.message
        if (-not $msg) { return }

        # Asistent: tool_use TaskCreate / TaskUpdate
        if ($msg.role -eq 'assistant' -and $msg.content) {
            foreach ($c in $msg.content) {
                if ($c.type -ne 'tool_use') { continue }
                if ($c.name -eq 'TaskCreate') {
                    $pendingCreate[$c.id] = $c.input.subject
                }
                elseif ($c.name -eq 'TaskUpdate') {
                    $id = "$($c.input.taskId)"
                    if (-not $tasks.ContainsKey($id)) { $tasks[$id] = @{ subject=''; status='pending'; updatedAt=$evt.timestamp } }
                    if ($c.input.status)  { $tasks[$id].status = $c.input.status }
                    if ($c.input.subject) { $tasks[$id].subject = $c.input.subject }
                    $tasks[$id].updatedAt = $evt.timestamp
                }
            }
        }
        # User: tool_result pentru TaskCreate -> extrag id-ul din text ("Task #N created successfully: Subject")
        elseif ($msg.role -eq 'user' -and $msg.content) {
            foreach ($c in $msg.content) {
                if ($c.type -ne 'tool_result') { continue }
                $useId = $c.tool_use_id
                if (-not $pendingCreate.ContainsKey($useId)) { continue }
                $subject = $pendingCreate[$useId]
                $txt = ''
                if ($c.content -is [string]) { $txt = $c.content }
                elseif ($c.content) { $txt = ($c.content | ForEach-Object { $_.text }) -join "`n" }
                $m = [regex]::Match($txt, 'Task #(\d+) created')
                if ($m.Success) {
                    $id = $m.Groups[1].Value
                    $createIdx++
                    $createOrder[$id] = $createIdx
                    $tasks[$id] = @{ subject=$subject; status='pending'; updatedAt=$evt.timestamp }
                }
                $pendingCreate.Remove($useId) | Out-Null
            }
        }
    }

    # Filtrez task-urile non-deleted și aleg ce să afișez:
    # 1) cel mai recent in_progress, altfel
    # 2) cel mai recent completed, altfel
    # 3) cel mai recent pending
    $live = @($tasks.GetEnumerator() | Where-Object { $_.Value.status -ne 'deleted' })
    if ($live.Count -eq 0) { Write-Output $line; exit 0 }

    $inProgress = $live | Where-Object { $_.Value.status -eq 'in_progress' } | Sort-Object { $_.Value.updatedAt } -Descending | Select-Object -First 1
    $completed  = $live | Where-Object { $_.Value.status -eq 'completed' }  | Sort-Object { $_.Value.updatedAt } -Descending | Select-Object -First 1
    $pending    = $live | Where-Object { $_.Value.status -eq 'pending' }    | Sort-Object { $_.Value.updatedAt } -Descending | Select-Object -First 1

    $openCount = (@($live | Where-Object { $_.Value.status -in 'in_progress','pending' })).Count
    $doneCount = (@($live | Where-Object { $_.Value.status -eq 'completed' })).Count

    if ($inProgress) {
        $sub = $inProgress.Value.subject
        $icon = [char]0x23F5  # ▶
        $line = "[$alias] $icon $sub  -  $openCount open / $doneCount done"
    } elseif ($completed) {
        $sub = $completed.Value.subject
        $icon = [char]0x2713  # ✓
        $line = "[$alias] $icon (ult) $sub  -  $openCount open / $doneCount done"
    } elseif ($pending) {
        $sub = $pending.Value.subject
        $line = "[$alias] o $sub  -  $openCount open / $doneCount done"
    }

    # Limitez lungimea
    if ($line.Length -gt 140) { $line = $line.Substring(0, 137) + '...' }
}
catch {
    $line = "[$alias] (statusline err: $($_.Exception.Message.Substring(0,[Math]::Min(60,$_.Exception.Message.Length))))"
}

Write-Output $line
