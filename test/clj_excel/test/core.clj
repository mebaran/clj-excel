(ns clj-excel.test.core
  (:use [clj-excel.core])
  (:use [clojure.test])
  (:import [java.io ByteArrayInputStream ByteArrayOutputStream]
           [org.apache.poi.ss.usermodel WorkbookFactory DateUtil Font]))

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
      (lazy-workbook #(lazy-sheet % cell-fn)) (postproc-wb)))

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
(defn val-link-map [cell]
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

(defn font-info [cell idx]
  (-> cell .getSheet .getWorkbook (.getFontAt (short idx)) bean
      (select-keys [:fontName :fontHeightInPoints :color])))

(defn extract-stylish [cell]
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
