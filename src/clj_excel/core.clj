(ns clj-excel.core
  (:use clojure.java.io)
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook])
  (:import [org.apache.poi.ss.usermodel Row Cell DateUtil WorkbookFactory]))

(def ^:dynamic *row-missing-policy* Row/CREATE_NULL_AS_BLANK)

(def ^:dynamic *data-formats* {:general 0 :number 1 :decimal 2 :comma 3 :accounting 4
                               :dollars 5 :red-neg 6 :cents 7 :dollars-red-neg 8
                               :percentage 9 :decimal-percentage 10 :scientific-notation 11
                               :short-ratio 12 :ratio 13
                               :date 14 :day-month-year 15 :day-month-name 16 :month-name-year 17
                               :hour-am-pm 18 :time-am-pm 1 :hour 20 :time 21 :datetime 22})

(defn data-format
  "Get dataformat by number or create new."
  [wb sformat]
  (if (keyword? sformat)
    (data-format wb (sformat *data-formats*))
    (-> wb .createDataFormat (.getFormat (if (number? sformat) (short sformat) sformat)))))

(defn cell-value
  "Return proper getter based on cell-value"
  ([cell] (cell-value cell (.getCellType cell)))
  ([cell cell-type]
      (condp = cell-type
        Cell/CELL_TYPE_BLANK nil
        Cell/CELL_TYPE_STRING (.getStringCellValue cell)
        Cell/CELL_TYPE_NUMERIC (if (DateUtil/isCellDateFormatted cell)
                                 (.getDateCellValue cell)
                                 (.getNumericCellValue cell))
        Cell/CELL_TYPE_BOOLEAN (.getBooleanCellValue cell)
        Cell/CELL_TYPE_FORMULA {:formula (.getCellFormula cell)}
        Cell/CELL_TYPE_ERROR {:error (.getErrorCellValue cell)}
        :unsupported)))

(defn workbook
  "Create or open new excel workbook."
  ([] (new XSSFWorkbook))
  ([input] (WorkbookFactory/create (input-stream input))))

(defn sheets
  "Get map of sheets."
  [wb] (zipmap (map #(.getSheetName %1) wb) (seq wb)))

(defn rows
  "Return rows from sheet as seq."
  [sheet] (seq sheet))

(defn cells
  "Return seq of cells from row"
  [row] (seq row))

(defn values
  "Return cells from sheet as seq."
  [row] (map cell-value (cells row)))

(defn lazy-sheet
  "Lazy seq of seq representing rows and cells."
  [sheet]
  (map #(map values %1) sheet))

(defn lazy-workbook
  "Lazy workbook report."
  [wb]
  (zipmap (map #(.getSheetName %1) wb) (lazy-sheet wb)))

(defn get-cell
  "Sell cell within row"
  ([row col] (.getCell row col))
  ([sheet row col] (get-cell (or (.getRow sheet row) (.createRow sheet row)) col)))

(defn coerce
  "Coerce cell for Java typing."
  [v]
  (cond
   (number? v) (double v)
   (or (symbol? v) (keyword? v)) (str v)
   :else v))

(defn set-cell
  "Set cell at specified location with value."
  ([cell value] (.setCellValue cell (coerce value)))
  ([row col value] (set-cell (or (get-cell row col) (.createCell row col)) value))
  ([sheet row col value] (set-cell (or (.getRow sheet row) (.createRow sheet row)) col value)))

(defn merge-rows
  "Add rows at end of sheet."
  [sheet start rows step]
  (doall
   (map
    (fn [rownum vals] (doall (map #(set-cell sheet rownum %1 %2) (iterate inc 0) vals)))
    (range start (+ start (count rows)))
    rows)))

(defn build-sheet
  "Build sheet from seq of seq (representing cells in row of rows)."
  [wb sheetname rows]
  (let [sheet (if sheetname
                (.createSheet wb sheetname)
                (.createSheet wb))]
    (merge-rows sheet 0 rows)))

(defn build-workbook
  "Build workbook from map of sheet names to multi dimensional seqs (ie a seq of seq)."
  [wb-map]
  (let [wb (workbook)]
    (doseq [[sheetname rows] wb-map]
      (build-sheet wb (str sheetname) rows))
    wb))

(defn save
  "Write worksheet to output-stream as coerced by OutputStream."
  [wb]
  (.save wb (output-stream wb)))