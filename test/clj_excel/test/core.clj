(ns clj-excel.test.core
  (:use [clj-excel.core])
  (:use [clojure.test])
  (:import [java.io ByteArrayInputStream ByteArrayOutputStream]
           [org.apache.poi.ss.usermodel WorkbookFactory DateUtil]))

;; restore data to nested vecs instead of seqs
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

(defn do-roundtrip [data mode]
  (-> (wb-from-data data mode) (save-load-cycle) (lazy-workbook) (postproc-wb)))

;; compare the data to the original
(defn valid-workbook-roundtrip? [data mode]
  (= data (do-roundtrip data mode)))


;; just numbers (doubles in poi); note: rows need not have equal length
(def trivial-input {"one" [[1.0] [2.0 3.0] [4.0 5.0 6.0]]})

(deftest roundtrip-trivial
  (is (valid-workbook-roundtrip? trivial-input :xssf))
  (is (valid-workbook-roundtrip? trivial-input :hssf)))


;; FIXME: Dates need hand holding :-(
(defn fix-date [e] [[(DateUtil/getJavaDate (ffirst e))]])
(defn fixed-roundtrip? [data mode]
  (-> (do-roundtrip data mode) (update-in ["four"] fix-date)
      (= data)))

(def now (java.util.Date.))
;; multiple sheets with different cell types
(def many-sheets {"one"   [[1.0]]   "two"  [["hello"]]
                  "three" [[false]] "four" [[now]]})

(deftest roundtrip-many
  (is (fixed-roundtrip? many-sheets :xssf))
  (is (fixed-roundtrip? many-sheets :hssf)))