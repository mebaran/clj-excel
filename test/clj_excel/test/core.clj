(ns clj-excel.test.core
  (:use [clj-excel.core])
  (:use [clojure.test])
  (:require [clojure.java.io :as io])
  (:import [java.io ByteArrayInputStream ByteArrayOutputStream]
           [org.apache.poi.ss.usermodel WorkbookFactory DateUtil Font Cell]))

;; restore data to nested vecs instead of seqs; equality test
(defn postproc-wb [m]
  (->> (for [[k v] m]
         [k (vec (map vec v))])
       (into {})))

;; build a workbook
(defn wb-from-data [data & opts]
  (let [as-set (set opts)
        wb     (cond (contains? as-set :hssf) (workbook-hssf)
                     :else                    (workbook-xssf))]
    (build-workbook wb data)))

;; save & reload
(defn save-load-cycle [wb]
  (let [os (ByteArrayOutputStream.)]
    (save wb os)
    (WorkbookFactory/create (ByteArrayInputStream. (.toByteArray os)))))

(defn do-roundtrip [data mode cell-fn]
  (-> (wb-from-data data mode) (save-load-cycle)
      (lazy-workbook #(lazy-sheet % :cell-fn cell-fn))
      (postproc-wb)))

;; compare the data to the original
(defn valid-workbook-roundtrip?
  ([data mode] (= data (do-roundtrip data mode cell-value)))
  ([data mode cell-fn] (= data (do-roundtrip data mode cell-fn))))

;; just numbers (doubles in poi); note: rows need not have equal length
(def trivial-input {"one" [[1.0] [2.0 3.0] [4.0 5.0 6.0]]})

(deftest roundtrip-trivial
  (is (valid-workbook-roundtrip? trivial-input :xssf))
  (is (valid-workbook-roundtrip? trivial-input :hssf)))


;; FIXME: Dates need hand holding :-(
(defn fix-date [e] [[(DateUtil/getJavaDate (ffirst e))]])
(defn fixed-roundtrip? [data mode]
  (-> (do-roundtrip data mode cell-value) (update-in ["four"] fix-date)
      (= data)))

(def now (java.util.Date.))
;; multiple sheets with different cell types
(def many-sheets {"one"   [[1.0]]   "two"  [["hello"]]
                  "three" [[false]] "four" [[now]]
                  "five"  [[nil]]})

(deftest roundtrip-many
  (is (fixed-roundtrip? many-sheets :xssf))
  (is (fixed-roundtrip? many-sheets :hssf)))

;; setting a map-typed object: value & hyperlink
(def url-link-input {"a" [[{:value "example.com" :link-url "http://www.example.com/"}]]})
(defn val-link-map [^Cell cell]
  {:value (cell-value cell) :link-url (.getAddress (.getHyperlink cell))})

(deftest cell-url-link
  (doseq [t [:hssf :xssf]]
    (is (valid-workbook-roundtrip? url-link-input t val-link-map))))

;; verify the fontspec api works
(def font-test-data
  [{:in {:font "Courier New" :size 16 :bold true}
    :out {:fontName "Courier New" :fontHeightInPoints 16
          :boldweight Font/BOLDWEIGHT_BOLD} }
   {:in {:font "Arial" :size 12 :italic true :color (color-indices :red) :strikeout true}
    :out {:fontName "Arial" :fontHeightInPoints 12 :italic true
          :color (color-indices :red) :strikeout true}}
   {:in {:underline :single}
    :out {:underline (underline-indices :single)}}])

(deftest fontspec-api
  (let [wb (workbook-hssf)]
    (doseq [{in :in out :out} font-test-data]
      (is (= (select-keys (bean (font wb in)) (keys out))
             out)))))

;; data format can be set by keyword or format string
(deftest dataformat-api
  (let [wb (workbook-hssf)]
    (is (= ((bean (create-cell-style wb :format :date)) :dataFormat)
           (data-formats :date)))
    (is (= ((bean (create-cell-style wb :format "yyyy-mm-dd")) :dataFormatString)
           "yyyy-mm-dd"))))

(deftest border-style-api
  (let [wb (workbook-hssf)]
    ;; all to the same type
    (is (= (select-keys (bean (create-cell-style wb :border :medium-dashed))
                        [:borderTop :borderRight :borderBottom :borderLeft])
           {:borderLeft 8, :borderBottom 8, :borderRight 8, :borderTop 8}))
    ;; grouped
    (is (= (select-keys (bean (create-cell-style wb :border [:none :medium]))
                        [:borderTop :borderRight :borderBottom :borderLeft])
           {:borderLeft 2, :borderBottom 0, :borderRight 2, :borderTop 0}))
    ;; individual styles
    (is (= (select-keys (bean (create-cell-style wb :border [:none :thin :medium :thick]))
                        [:borderTop :borderRight :borderBottom :borderLeft])
           {:borderLeft 5, :borderBottom 2, :borderRight 1, :borderTop 0}))))


;; playing with cell styles
;; note: hyperlink-cell have unreadable color defaults; you better set those
(def stylish-test-data
   {"foo" [[{:value "Hello world" :font {:font "Courier New" :size 16 :color :blue}
            :foreground-color :maroon :pattern :solid-foreground}]]
    "bar" [[{:value "click me" :link-url "http://www.example.com/"
             :font {:color :black :font "Serif" :size 10}}]]})

(defn font-info [^Cell cell idx]
  (-> cell .getSheet .getWorkbook (.getFontAt (short idx)) bean
      (select-keys [:fontName :fontHeightInPoints :color])))

(defn extract-stylish [^Cell cell]
  (merge (hash-map :value (cell-value cell)
                   :style (select-keys (bean (.getCellStyle cell))
                                       [:fillPattern :fillForegroundColor])
                   :font (font-info cell (.getFontIndex (.getCellStyle cell))))
         (when-let [link (.getHyperlink cell)]
           {:link-url (.getAddress link)})))

;; note: needs explicit fonts; different defaults xls: Arial, xlsx: Colibri
(deftest stylish-test
  (let [expected {"bar"
                  [[{:style {:fillForegroundColor 64, :fillPattern 0},
                     :link-url "http://www.example.com/",
                     :font {:color 8, :fontHeightInPoints 10, :fontName "Serif"},
                     :value "click me"}]],
                  "foo"
                  [[{:style {:fillForegroundColor 25, :fillPattern 1},
                     :font {:color 12, :fontHeightInPoints 16, :fontName "Courier New"},
                     :value "Hello world"}]]}]
    (is (= (do-roundtrip stylish-test-data :hssf extract-stylish)
           expected))
    (is (= (do-roundtrip stylish-test-data :xssf extract-stylish)
           expected))))

(deftest cell-mutator-test
  (let [wb (wb-from-data {"sheet1" [[nil]]} :hssf)
        sheet (-> wb sheets first)
        ^Cell cell (get-cell sheet 0 0)]
    (testing "Setting a boolean"
      (cell-mutator cell true)
      (is (.getBooleanCellValue cell)))
    (testing "Setting a number"
      (cell-mutator cell 1)
      (is (= 1.0 (.getNumericCellValue cell))))
    (testing "Setting a string"
      (cell-mutator cell "foo")
      (is (= "foo" (.getStringCellValue cell))))
    (testing "Setting a keyword"
      (cell-mutator cell :foo)
      (is (= "foo" (.getStringCellValue cell))))
    (testing "Setting a date"
      (cell-mutator cell #inst "2013-09-11")
      (is (= #inst "2013-09-11" (.getDateCellValue cell))))
    (testing "Setting nil"
      (cell-mutator cell nil)
      (is (= Cell/CELL_TYPE_BLANK (.getCellType cell))))
    (testing "Setting a cell style"
      (let [cs (create-cell-style wb)]
        (cell-mutator cell {:style cs})
        (is (= cs (.getCellStyle cell)))))
    (testing "Setting a formula"
      (cell-mutator cell {:formula "A1"})
      (is (= "A1" (.getCellFormula cell))))))

(deftest row-seq-test
  (let [wb (workbook-hssf (io/resource "test-nil-cell-1.xls"))
        row (-> wb (.getSheetAt 0) second)]
    (testing "Default mode is logical"
      (is (= [1.0 nil 3.0] (row-seq row))))
    (testing "Mode logical"
      (is (= [1.0 nil 3.0] (row-seq row :mode :logical))))
    (testing "Mode physical"
      (is (= [1.0 3.0] (row-seq row :mode :physical))))))

(deftest lazy-sheet-test
  (let [wb (workbook-hssf (io/resource "test-nil-cell-1.xls"))
        sheet (.getSheetAt wb 0)]
    (testing "Default mode is logical"
      (is (= [["A" "B" "C"] [1.0 nil 3.0]] (lazy-sheet sheet))))
    (testing "Mode physical"
      (is (= [["A" "B" "C"] [1.0 3.0]] (lazy-sheet sheet :mode :physical))))))

(def comment-test-data
  {"foo" [[{:value "Hello world" :comment {:text "Lorem Ipsum" :width 10 :height 3}}]]})

(defn extract-comment [cell]
  (merge (hash-map :value (cell-value cell)
                   :comment (cell-comment cell))))

(deftest comment-test
  (let [expected {"foo"
                  [[{:value "Hello world"
                     :comment {:text "Lorem Ipsum"}}]]}]
    (is (= (do-roundtrip comment-test-data :hssf extract-comment)
           expected))
    (is (= (do-roundtrip comment-test-data :xssf extract-comment)
           expected))))
