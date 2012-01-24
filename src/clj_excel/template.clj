(ns clj-excel.template
  (:use clj-excel.core)
  (:import [org.apache.poi.ss.usermodel CellStyle]))

(def ^:dynamic *styles* {})

(defn cell-style
  "Lookup or return cell-style"
  [wb stylespec]
  (cond
   (isa? (type stylespec) CellStyle) stylespec
   :else (create-cell-style wb (apply concat stylespec))))

(defn cell-properties
  "Map of pertinent properties.  Includes row, col, value currently."
  [cell]
  {:col (.getColumnIndex cell)
   :row (.getRowIndex cell)
   :value (cell-value cell)})

(defn style-cells
  "Templater - takes two functions to load data and style workbooks.
  Takes a hash-map of names to style constructions a collection of
  criteria, and a function to produce the the data:
    -- Map of style names as keywords to style specs
    -- Map of keywords to filter criteria"
  [wb stylebook stylesheet cells]
  (let [style-dict (zipmap (keys stylebook) (map cell-style (vals stylebook)))
        style-filters (zipmap (keys stylesheet) (style-filter (vals stylesheet)))]
    ()))