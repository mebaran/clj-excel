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

(defn)

(defn style-cells
  "Templater - takes two functions to load data and style workbooks.
  Takes a hash-map of names to style constructions a collection of
  criteria, and a function to produce the the data"
  [style criteria cells]
  (fn [wb] (doall (map #(.setCellStyle %1 (cell-style wb style)) (filter criteria cells)))))