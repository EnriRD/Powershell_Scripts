Get-Content -Path  '.\search values.txt' `
    | ForEach-Object { $_.Insert(2,":").Insert(8,":").Insert(14,":") } `
    | Out-File -FilePath '.\output.txt' -Append