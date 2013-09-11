(ns clj-excel.core
  (:use clojure.java.io)
  (:import [clojure.lang Keyword APersistentMap])
  (:import java.util.Date)
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook])
  (:import [org.apache.poi.hssf.usermodel HSSFWorkbook])
  (:import [org.apache.poi.ss.usermodel Row Cell DateUtil WorkbookFactory CellStyle Font
            Hyperlink Workbook Sheet]))

(def ^:dynamic *row-missing-policy* Row/CREATE_NULL_AS_BLANK)

(def data-formats {:general 0 :number 1 :decimal 2 :comma 3 :accounting 4
                   :dollars 5 :red-neg 6 :cents 7 :dollars-red-neg 8
                   :percentage 9 :decimal-percentage 10 :scientific-notation 11
                   :short-ratio 12 :ratio 13
                   :date 14 :day-month-year 15 :day-month-name 16 :month-name-year 17
                   :hour-am-pm 18 :time-am-pm 1 :hour 20 :time 21 :datetime 22})

(def color-indices {:black 8, :violet 20, :blue-grey 54, :dark-yellow 19,
                    :automatic 64, :grey-50-percent 23, :light-cornflower-blue 31,
                    :light-turquoise 41, :white 9, :plum 61, :orange 53, :teal 21,
                    :olive-green 59, :red 10, :lavender 46, :light-orange 52, :brown 60,
                    :light-green 42, :light-yellow 43, :royal-blue 30, :gold 51,
                    :aqua 49, :coral 29, :light-blue 48, :blue 12, :cornflower-blue 24,
                    :dark-green 58, :grey-25-percent 22, :green 17, :lemon-chiffon 26,
                    :lime 50, :dark-blue 18, :sky-blue 40, :pale-blue 44, :dark-red 16,
                    :dark-teal 56, :grey-80-percent 63, :indigo 62, :sea-green 57, :pink 14,
                    :turquoise 15, :tan 47, :yellow 13, :orchid 28, :grey-40-percent 55,
                    :rose 45, :maroon 25, :bright-green 11})

(def underline-indices {:none 0 :single 1 :double 2 :single-accounting 33 :double-accounting 34})

;; Utility Constant Look Up ()

(defn constantize
  "Helper to read constants from constant like keywords within a class.  Reflection powered."
  [^Class klass kw]
  (.get (.getDeclaredField klass (-> kw name (.replace "-" "_") .toUpperCase)) Object))

(defn cell-style-constant
  ([kw prefix]
     (if (number? kw)
       (short kw)
       (short (constantize CellStyle (if prefix
                                       (str
                                        (name prefix) "-"
                                        (-> kw name
                                            (.replaceFirst (str (name prefix) "-") "")
                                            (.replaceFirst (str (name prefix) "_") "")
                                            (.replaceFirst (name prefix) "")))
                                       kw)))))
  ([kw] (cell-style-constant kw nil)))

;; Workbook and Style functions

(defn data-format
  "Get dataformat by number or create new."
  [^Workbook wb sformat]
  (cond
   (keyword? sformat) (data-format wb (sformat data-formats))
   (number? sformat) (short sformat)
   (string? sformat) (-> wb .getCreationHelper .createDataFormat (.getFormat ^String sformat))))

(defn set-border
  "Set borders, css order style.  Borders set CSS order."
  ([cs all] (set-border cs all all all all))
  ([cs caps sides] (set-border cs caps sides caps sides))
  ([^CellStyle cs top right bottom left] ;; CSS ordering
     (.setBorderTop cs (cell-style-constant top :border))
     (.setBorderRight cs (cell-style-constant right :border))
     (.setBorderBottom cs (cell-style-constant bottom :border))
     (.setBorderLeft cs (cell-style-constant left :border))))

(defn- col-idx [v]
  (short (if (keyword? v) (color-indices v) v)))

(defn font
  "Register font with "
  [^Workbook wb fontspec]
  (if (isa? (type fontspec) Font)
    fontspec
    (let [default-font (.getFontAt wb (short 0)) ;; First font is default
          boldweight (short (get fontspec :boldweight (if (:bold fontspec)
                                                        Font/BOLDWEIGHT_BOLD
                                                        Font/BOLDWEIGHT_NORMAL)))
          color (short (if-let [k (fontspec :color)]
                         (col-idx k)
                         (.getColor default-font)))
          size (short (get fontspec :size (.getFontHeightInPoints default-font)))
          name (str (get fontspec :font (.getFontName default-font)))
          italic (boolean (get fontspec :italic false))
          strikeout (boolean (get fontspec :strikeout false))
          typeoffset (short (get fontspec :typeoffset 0))
          underline (byte (if-let [k (fontspec :underline)]
                            (if (keyword? k) (underline-indices k) k)
                            (.getUnderline default-font)))]
      (or
       (.findFont wb boldweight size color name italic strikeout typeoffset underline)
       (doto (.createFont wb)
         (.setBoldweight boldweight)
         (.setColor color)
         (.setFontName name)
         (.setItalic italic)
         (.setStrikeout strikeout)
         (.setFontHeightInPoints size)
         (.setUnderline underline))))))

(defn create-cell-style
  "Create style for workbook"
  [^Workbook wb & {format :format alignment :alignment border :border fontspec :font
         bg-color :background-color fg-color :foreground-color pattern :pattern}]
  (let [cell-style (.createCellStyle wb)]
    (if fontspec (.setFont cell-style (font wb fontspec)))
    (if format (.setDataFormat cell-style (data-format wb format)))
    (if alignment (.setAlignment cell-style (cell-style-constant alignment :align)))
    (if border (if (coll? border)
                 (apply set-border cell-style border)
                 (set-border cell-style border)))
    (if fg-color (.setFillForegroundColor cell-style (col-idx fg-color)))
    (if bg-color (.setFillBackgroundColor cell-style (col-idx bg-color)))
    (if pattern  (.setFillPattern cell-style (cell-style-constant pattern)))
    cell-style))

;; extract the sub-map of options supported by create-cell-style
(defn- get-style-attributes [m]
  (select-keys m [:format :alignment :border :font :background-color :foreground-color :pattern]))

(defprotocol StyleCache
  (build-style [this cell-data]))

;; iterate a nested sheet-data seq and create cell styles
(defn create-sheet-data-style [cache data]
  (for [row data]
    (for [col row]
      (if (and (map? col) (not (empty? (get-style-attributes col))))
        (assoc col :style (build-style cache (get-style-attributes col)))
        col))))

(defn caching-style-builder [wb]
  (let [cache (atom {})]
    (reify StyleCache
      (build-style [_ style-key]
        (if-let [style (get @cache style-key)]
          style
          (let [style (apply create-cell-style wb (reduce #(conj %1 (first %2) (second %2)) [] style-key))]
            (swap! cache assoc style-key style)
            style))))))

;; Reading functions

(defn cell-value
  "Return proper getter based on cell-value"
  ([^Cell cell] (cell-value cell (.getCellType cell)))
  ([^Cell cell cell-type]
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

(defn ^Workbook workbook-xssf
  "Create or open new excel workbook. Defaults to xlsx format."
  ([] (new XSSFWorkbook))
  ([input] (WorkbookFactory/create (input-stream input))))

(defn ^Workbook workbook-hssf
  "Create or open new excel workbook. Defaults to xls format."
  ([] (new HSSFWorkbook))
  ([input] (WorkbookFactory/create (input-stream input))))

(defn sheets
  "Get seq of sheets."
  [^Workbook wb] (map #(.getSheetAt wb %1) (range 0 (.getNumberOfSheets wb))))

(defn rows
  "Return rows from sheet as seq.  Simple seq cast via Iterable implementation."
  [sheet] (seq sheet))

(defn cells
  "Return seq of cells from row.  Simpel seq cast via Iterable implementation." 
  [row] (seq row))

(defn values
  "Return cells from sheet as seq."
  [row] (map cell-value (cells row)))

(defn row-seq
  "Returns a lazy seq of cells of row.
  
  Options:
    :cell-fn function called on each cell, defaults to cell-value
    :mode    either :logical (default) or :physical
  
  Modes:
    :logical  returns all cells even if they are blank
    :physical returns only the physically defined cells"
  {:arglists '([row & opts])}
  [^Row row & {:keys [cell-fn mode] :or {cell-fn cell-value mode :logical}}] 
  (condp = mode
    :logical (map #(when-let [cell (.getCell row %)] (cell-fn cell)) (range 0 (.getLastCellNum row)))
    :physical (map cell-fn row)
    (throw (ex-info (str "Unknown mode " mode) {:mode mode}))))

(defn lazy-sheet
  "Lazy seq of seqs representing rows and cells of sheet.
  
  Options:
    :cell-fn function called on each cell, defaults to cell-value
    :mode    either :logical (default) or :physical
  
  Modes:
    :logical  returns all cells even if they are blank
    :physical returns only the physically defined cells"
  {:arglists '([sheet & opts])}
  [sheet & {:keys [cell-fn mode] :or {cell-fn cell-value mode :logical}}] 
  (map #(row-seq % :cell-fn cell-fn :mode mode) sheet))

(defn sheet-names
  [^Workbook wb]
  (->> (.getNumberOfSheets wb) (range) (map #(.getSheetName wb %))))

(defn lazy-workbook
  "Lazy workbook report."
  ([wb] (lazy-workbook wb lazy-sheet))
  ([wb sheet-fn] (zipmap (sheet-names wb) (map sheet-fn (sheets wb)))))

(defn get-cell
  "Sell cell within row"
  ([^Row row col] (.getCell row col))
  ([^Sheet sheet row col] (get-cell (or (.getRow sheet row) (.createRow sheet row)) col)))

;; Writing Functions

(defn- get-link-type [m]
  (some #{:link-url :link-email :link-document :like-file} (keys m)))

(defn create-link [^Cell cell kw link-to]
  (let [link-type (constantize org.apache.poi.common.usermodel.Hyperlink kw)
        link      (-> cell .getSheet .getWorkbook .getCreationHelper
                      (.createHyperlink link-type))]
    (.setAddress link link-to)
    (.setHyperlink cell link)))

(defmulti cell-mutator (fn [cell val] (class val)))
(defmethod cell-mutator Boolean [^Cell cell ^Boolean b] (.setCellValue cell b))
(defmethod cell-mutator Number [^Cell cell n] (.setCellValue cell (double n)))
(defmethod cell-mutator String [^Cell cell ^String s] (.setCellValue cell s))
(defmethod cell-mutator Keyword [^Cell cell kw] (.setCellValue cell (name kw)))
(defmethod cell-mutator Date [^Cell cell ^Date date] (.setCellValue cell date))
(defmethod cell-mutator nil [^Cell cell null] (.setCellType cell Cell/CELL_TYPE_BLANK))
(defmethod cell-mutator APersistentMap [^Cell cell m]
  (cell-mutator cell (m :value))
  (when-let [link-key (get-link-type m)]
    (create-link cell link-key (m link-key)))
  (when-let [style (m :style)]
    (.setCellStyle cell style))
  (when-let [formula (m :formula)]
    (.setCellFormula cell formula)))

(defn set-cell
  "Set cell at specified location with value."
  ([cell value] (cell-mutator cell value))
  ([^Row row col value] (set-cell (or (get-cell row col) (.createCell row col)) value))
  ([^Sheet sheet row col value] (set-cell (or (.getRow sheet row) (.createRow sheet row)) col value)))

(defn merge-rows
  "Add rows at end of sheet."
  [sheet start rows]
  (doall
   (map
    (fn [rownum vals] (doall (map #(set-cell sheet rownum %1 %2) (iterate inc 0) vals)))
    (range start (+ start (count rows)))
    rows)))

(defn build-sheet
  "Build sheet from seq of seq (representing cells in row of rows)."
  [^Workbook wb sheetname rows]
  (let [sheet (if sheetname
                (.createSheet wb sheetname)
                (.createSheet wb))]
    (merge-rows sheet 0 rows)))

(defn build-workbook
  "Build workbook from map of sheet names to multi dimensional seqs (ie a seq of seq)."
  ([wb wb-map]
     (let [cache (caching-style-builder wb)]
       (doseq [[sheetname rows] wb-map]
         (build-sheet wb (str sheetname) (create-sheet-data-style cache rows)))
       wb))
  ([wb-map] (build-workbook (workbook-xssf) wb-map)))

(defn save
  "Write workbook to output-stream as coerced by OutputStream."
  [^Workbook wb path]
  (with-open [out (output-stream path)]
    (.write wb out)))
