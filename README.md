# Project Title

CreateProcessFromIShellDispatchInvoke

## Description

COM hooking as well as alternate ways to execute binaries from COM via IDispatch interface.
BLOG POST HERE: https://mohamed-fakroud.gitbook.io/red-teamings-dojo/windows-internals/playing-around-com-objects-part-1

## Getting Started

### Executing program
```
> gcc .\comfunc.c -o com.exe -lole32 -luuid -loleaut32
> .\com.exe C:\Windows\System32\calc.exe (The absolute path of your executable)
```

## Author
[@t3nb3w](https://twitter.com/T3nb3w)

## Acknowledgments
* [@MrUn1k0d3r](https://twitter.com/MrUn1k0d3r)
