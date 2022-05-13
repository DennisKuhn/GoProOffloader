# GoProOffloader
Powershell script to automaticlly offload go pro files.

## Minimal path to awesomeness
To get started:
- Download the script
- configure the destination folder variable at the top
- Run the script once as administrator to enable logging to the system log
- create task with the task scheduler
  - trigger on login or startup.
  - Add an action
     - Program/script: C:\Program Files\PowerShell\7\pwsh.exe
     - Add arguments option: -File "REPLACE-WITH-YOUR-PATH\GoPro-Offload.ps1"
