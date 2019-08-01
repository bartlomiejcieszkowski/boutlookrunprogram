# bOutlookRunProgram
simple extension for outlook where you can write custom rules in xml

for sample xml see [this](XMLFile1.xml)

## Known issues with vsto
It is possible that when one uninstall a vsto plugin for outlook it will prompt that the extension is already installed even though we just uninstalled it. 

Possible workaround for this issue is to run:
```
rundll32 dfshim CleanOnlineAppCache
```
