(ns clj-excel.test.core
  (:use [clj-excel.core])
  (:use [clojure.test])
  (:import [java.io ByteArrayInputStream ByteArrayOutputStream]
           [org.apache.poi.ss.usermodel WorkbookFactory]))

;; just numbers (doubles in poi); note: rows need not have equal length
(def trivial-input {"one" [[1.0] [2.0 3.0] [4.0 5.0 6.0]]})

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

;; compare the data to the original
(defn valid-workbook-roundtrip? [data mode]
  (-> (wb-from-data data mode) (save-load-cycle) (lazy-workbook) (postproc-wb)
      (= data)))

(deftest roundtrip-trivial
  (is (valid-workbook-roundtrip? trivial-input :xssf))
  (is (valid-workbook-roundtrip? trivial-input :hssf)))