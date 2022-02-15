# InspectAppsUtils

InspectAppsUtils is VBA Modules that complement Microsoft Office's document inspection.

## Description

In Microsoft Office, the document inspection is a good tool for preventing information leaks. But we often think "OK, I understood that there are some hidden items, but where are?", "Don't remove, SHOW it." and so on.
InspectAppsUtils provides functions: list them, visualize them, and delete them with inspectee information.

## Usage

Import class file and construct.

```vb
Dim inspectUtil As InspectWorkbookUtils
Set inspectUtil = New InspectWorkbookUtils

Call inspectUtil.Initialize(ThisWorkbook) 
```

The Initialize method must be called earlier than any other method.

If you want to know about each method, please read the help with object browser.

## License

MIT

Copyright (c) 2022 telitas

See the LICENSE.txt file or https://opensource.org/licenses/mit-license.php for details.
