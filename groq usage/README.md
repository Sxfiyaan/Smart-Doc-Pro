# Groq usage breakdown

SmartDoc records one JSON object in `requests.jsonl` for each Groq request made
for file inspection or a command/tool loop. The log includes provider-reported
prompt, completion, total, cached, and reasoning token counts when Groq returns
them. It also estimates prompt size by component: system instructions, previous
SmartDoc chat history, current file summaries and command, tool definitions, and
messages accumulated during the tool loop.

Provider totals are exact. Component token counts are estimates based on text
length (characters divided by four); Groq does not provide exact per-message
token attribution in the response usage object.

The log intentionally excludes prompt text, commands, file contents, and file
names. Inspect it with any JSON Lines reader or PowerShell:

```powershell
Get-Content '.\groq usage\requests.jsonl' | ForEach-Object { $_ | ConvertFrom-Json }
```

The live web app also shows per-request totals and the approximate component
breakdown after inspection and command runs. Start a new request after restarting
the backend to begin collecting usage records.
