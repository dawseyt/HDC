$logdir = "test_dir"
New-Item -ItemType Directory -Force -Path $logdir > $null

for ($i=0; $i -lt 1000; $i++) {
    New-Item -ItemType File -Force -Path "$logdir\UnlockLog_$i.csv" > $null
    New-Item -ItemType File -Force -Path "$logdir\OtherLog_$i.txt" > $null
}

$escaped = $logdir.Replace('[', '`[').Replace(']', '`]')

Measure-Command {
    Get-ChildItem -LiteralPath $logdir | Where-Object { $_.Name -like "UnlockLog_*.csv" } > $null
}

Measure-Command {
    Get-ChildItem -Path $escaped -Filter "UnlockLog_*.csv" > $null
}
