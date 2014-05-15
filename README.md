# clj-excel: The Beauty And The Beast

The goal is to give you a carefree experience while using
[Apache POI][poi-home] from Clojure.

Please note that the API isn't stable yet!

[![Build Status](https://travis-ci.org/undernorthernsky/clj-excel.png)](https://travis-ci.org/undernorthernsky/clj-excel)

[poi-home]: http://poi.apache.org/ "The Java API for Microsoft Documents"

## Usage

Saving and loading data:

```clojure
(use 'clj-excel.core)
(-> (build-workbook (workbook-hssf) {"Numbers" [[1] [2 3] [4 5 6]]})
    (save "numbers.xls"))

(lazy-workbook (workbook-hssf "numbers.xls"))
; {"Numbers" ((1.0) (2.0 3.0) (4.0 5.0 6.0))}
```

Cell values can be any type supported by POI (`boolean`, `double`,
`String`, `Date`, ...; see [setCellValue(...)][poi-types]).

They may also be maps; this enables styling:

```clojure
(def a-cell-value
  {:value "world" :alignment :center
   :border [:none :thin :dashed :thin]
   :foreground-color :grey-25-percent :pattern :solid-foreground
   :font {:color :blue :underline :single :italic true
          :size 12 :font "Arial"}})

(-> (build-workbook (workbook-hssf) {"hello" [[a-cell-value]]})
    (save "hello-world.xls"))
```

Creating links:

```clojure
;; just the data
{"a" [[{:value "foo" :link-document "b!A1"}]]
 "b" [[{:value "example.com" :link-url "http://www.example.com/"}]]}
```

Creating comments:

```clojure
{"a" [[{:value "foo" :comment {:text "Lorem Ipsum" :width 4 :height 2}}]]})
```

[poi-types]: http://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Cell.html

## TODO

* more concise styling; inheriting from declared styles?
* support for images
* autoresize columns
* convienent in-document links

## Relevant links

Depending on your needs:

* [Guide][poi-guide]: How to get things done ...
* [API docs][poi-jdoc]: ... in a little more detail
* [the source][poi-src]: ... when everything else fails

[poi-guide]: http://poi.apache.org/spreadsheet/quick-guide.html "Busy Developers' Guide ..."
[poi-jdoc]: http://poi.apache.org/apidocs/index.html "POI Javadoc"
[poi-src]: http://svn.apache.org/repos/asf/poi/trunk/ "POI source code"

## License

Copyright (C) 2012 FIXME

Distributed under the Eclipse Public License, the same as Clojure.
