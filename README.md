# powershell-scripts
Some of my powershell scripts

---

### Unrecycle.ps1
```
Usage: .\Unrecycle.ps1 [-ItemNameToRestore <RecycleBinItemName>] [-Verb <VerbName>]
Examples:
  (Check recycle bin contents)      .\Unrecycle.ps1
  (Restore an item using default)   .\Unrecycle.ps1 -ItemNameToRestore 'secret_plans.docx'
  (Use a different verb)            .\Unrecycle.ps1 -ItemNameToRestore 'secret_plans.docx' -Verb 'undelete'

```
Checks the recycle bin for any contents. If there's anything in there, the script will mention it.
If the 'restore' verb doesn't work to restore the item, the script will tell you which verbs are available.
