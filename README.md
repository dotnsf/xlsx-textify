# XLSX Textify


## Overview

Convert Excel(XLSX) file to JSON text


## How to textify

- Convert this styled Excel sheet...

```
+-----+-----+-----+
| AAA | BBB | CCC |
+-----+-----+-----+
| dd  | eee | f   |
+-----+-----+-----+
| ggg | h   | iii |
+-----+-----+-----+
| j   | kkk | lll |
+-----+-----+-----+
| mm  | nn  | oo  |
+-----+-----+-----+
```

- into this JSON array:

```
[
  { "AAA": "dd", "BBB": "eee", "CCC": "f" },
  { "AAA": "ggg", "BBB": "h", "CCC": "iii" },
  { "AAA": "j", "BBB": "kkk", "CCC": "lll" },
  { "AAA": "mm", "BBB": "nn", "CCC": "oo" }
]
```


## Licensing

This code is licensed under MIT.


## Copyright

2024  [K.Kimura @ Juge.Me](https://github.com/dotnsf) all rights reserved.
