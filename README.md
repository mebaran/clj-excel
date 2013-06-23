# clj-excel: The Beauty And The Beast

The goal is to give you a carefree experience while using
[Apache POI][poi-home] from Clojure.

Please note that the API isn't stable yet!

[poi-home]: http://poi.apache.org/ "The Java API for Microsoft

## Usage

Saving and loading data:

```clojure
(use 'clj-excel.core)
(-> (build-workbook (workbook-hssf) {"Numbers" [[1] [2 3] [4 5 6]]})
    (save "numbers.xls"))

(lazy-workbook (workbook-hssf "numbers.xls"))
; {"Numbers" ((1.0) (2.0 3.0) (4.0 5.0 6.0))}
```

## License

Copyright (C) 2012 FIXME

Distributed under the Eclipse Public License, the same as Clojure.
