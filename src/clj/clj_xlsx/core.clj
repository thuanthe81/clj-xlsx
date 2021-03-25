(ns clj-xlsx.core
  (:require [clojure.string :as str])
  (:import (org.apache.poi.xssf.usermodel XSSFCell
                                          XSSFCellStyle
                                          XSSFFont
                                          XSSFColor
                                          XSSFRow
                                          XSSFSheet
                                          XSSFWorkbook
                                          XSSFFormulaEvaluator)
           (java.io FileInputStream
                    FileOutputStream
                    InputStream)
           (org.apache.poi.ss.usermodel Cell
                                        CellType
                                        HorizontalAlignment
                                        VerticalAlignment WorkbookFactory Workbook Sheet CellValue DateUtil FormulaError)
           (java.net URL)
           (java.util Date)))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;; READING XLSX ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;; Refer xlsx.schema for reading/writing excel ;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



(defn xsssf-color->css-color
  [^XSSFColor c]
  (some-> c
          .getARGBHex
          (subs 2)
          .toLowerCase
          (#(str "#" %))))

(defn css-color->xsssf-format
  [^String str-color]
  (str/replace str-color #"#" ""))

(defn ^XSSFColor xsssf-color-fill-by-css
  [^XSSFColor c css-color]
  (->> css-color
       css-color->xsssf-format
       (.setARGBHex c)))

(defmulti read-formula-value (fn [^CellValue cv date-format?]
                               (.getCellType cv)))
(defmethod read-formula-value CellType/BOOLEAN
  [^CellValue cv _]
  (.getBooleanValue cv))
(defmethod read-formula-value CellType/STRING
  [^CellValue cv _]
  (.getStringValue cv))
(defmethod read-formula-value CellType/NUMERIC
  [^CellValue cv date-format?]
  (if date-format?
    (DateUtil/getJavaDate (.getNumberValue cv))
    (.getNumberValue cv)))
(defmethod read-formula-value CellType/ERROR
  [^CellValue cv _]
  (keyword (.name (FormulaError/forInt (.getErrorValue cv)))))


(defmulti read-cell-value (fn [^Cell cell]
                            (some-> cell .getCellType)))
(defmethod read-cell-value CellType/STRING
  [^Cell cell]
  {:v (.getStringCellValue cell)})
(defmethod read-cell-value CellType/BOOLEAN
  [^Cell cell]
  {:v (.getBooleanCellValue cell)})
(defmethod read-cell-value CellType/NUMERIC
  [^Cell cell]
  {:v (if (DateUtil/isCellDateFormatted cell)
        (.getDateCellValue cell)
        (.getNumericCellValue cell))})
(defmethod read-cell-value CellType/FORMULA
  [^Cell cell]
  (let [f         (.getCellFormula cell)
        evaluator (.. cell getSheet getWorkbook
                      getCreationHelper createFormulaEvaluator)
        cv        (.evaluate evaluator cell)]
    {:f f
     :v (if (and (= CellType/NUMERIC (.getCellType cv))
                 (DateUtil/isCellDateFormatted cell))
          (.getDateCellValue cell)
          (read-formula-value cv false))}))
(defmethod read-cell-value CellType/ERROR
  [^Cell cell]
  {:v (keyword (.name (FormulaError/forInt (.getErrorCellValue cell))))})
(defmethod read-cell-value :default
  [_]
  nil)

(defn horizontal-alignment->keyword
  [alignment]
  (condp = alignment
    HorizontalAlignment/LEFT :left
    HorizontalAlignment/RIGHT :right
    HorizontalAlignment/CENTER :center
    HorizontalAlignment/JUSTIFY :justify
    nil))

(defn keyword->horizontal-alignment
  [key]
  (condp = key
    :left HorizontalAlignment/LEFT
    :right HorizontalAlignment/RIGHT
    :center HorizontalAlignment/CENTER
    :justify HorizontalAlignment/JUSTIFY
    nil))


(defn vertical-alignment->keyword
  [alignment]
  (condp = alignment
    VerticalAlignment/TOP :top
    VerticalAlignment/BOTTOM :bottom
    VerticalAlignment/CENTER :center
    VerticalAlignment/JUSTIFY :justify
    nil))

(defn keyword->vertical-alignment
  [key]
  (condp = key
    :top VerticalAlignment/TOP
    :bottom VerticalAlignment/BOTTOM
    :center VerticalAlignment/CENTER
    :justify VerticalAlignment/JUSTIFY
    nil))


(defn read-font
  [^XSSFFont font]
  (when font
    (merge {:name   (.getFontName font)
            :size   (.getFontHeightInPoints font)}
           (when-let [color (some->> font
                                     .getXSSFColor
                                     xsssf-color->css-color)]
             {:color color})
           (when (.getBold font)
             {:bold true})
           (when (.getItalic font)
             {:italic true}))))


(defn read-cell
  [^Cell cell]
  (let [v (read-cell-value cell)]
    (when (:v v)
      (merge v
             {:s (let [cell-s (.getCellStyle cell)
                       font   (.getFont ^XSSFCellStyle cell-s)]
                   (merge {:fmt {:n (.getDataFormat cell-s)}}
                          (when (.getWrapText cell-s)
                            {:wrap-text true})
                          (when-let [alignment (some-> (.getAlignment cell-s)
                                                       horizontal-alignment->keyword)]
                            {:horizontal-alignment alignment})
                          (when-let [alignment (some-> (.getVerticalAlignment cell-s)
                                                       vertical-alignment->keyword)]
                            {:vertical-alignment alignment})
                          (some->> (.getFillBackgroundColorColor cell-s)
                                   xsssf-color->css-color
                                   (hash-map :background-color))
                          (when font
                            {:font (read-font font)})))}))))


(defn read-row
  [^XSSFSheet sh row-i]
  (let [row       (.getRow sh row-i)
        num-cells (.getLastCellNum row)]
    (->> num-cells
         range
         (mapv (comp read-cell
                     #(.getCell row %))))))


(defn read-sheet
  [^XSSFSheet sheet]
  (let [num-rows (.getLastRowNum sheet)]
    {:name (.getSheetName sheet)
     :rows (->> num-rows
                range
                (mapv (partial read-row sheet)))}))


(defmulti load-workbook class)
(defmethod load-workbook String
  [^String filename]
  (with-open [stream (FileInputStream. filename)]
    (load-workbook stream)))
(defmethod load-workbook InputStream
  [^InputStream stream]
  (WorkbookFactory/create stream))
(defmethod load-workbook URL
  [^URL url]
  (with-open [stream (.openStream url)]
    (load-workbook stream)))

(defn resource->URL
  [^String resource]
  (clojure.java.io/resource resource))


(defn read-workbook
  [from]
  (let [wb (load-workbook from)]
    (->> (.getNumberOfSheets wb)
         range
         (mapv (comp read-sheet
                     #(.getSheetAt wb %))))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;; WRITING XLSX ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defn ^XSSFFont fill-font
  [^XSSFFont cs-f {:keys [name size bold italic]}]
  (cond-> nil
          name ((fn [_] (.setFontName cs-f name)))
          size ((fn [_] (.setFontHeightInPoints cs-f ^Short size)))
          bold ((fn [_] (.setBold cs-f bold)))
          italic ((fn [_] (.setItalic cs-f italic)))))


(defn fill-style
  [^Workbook wb ^XSSFCellStyle cs {:keys [font fmt background-color wrap-text horizontal-alignment vertical-alignment]}]
  (when background-color
    (-> cs
        .getFillBackgroundXSSFColor
        ((fn [^XSSFColor color]
           (let [color (or color
                           (XSSFColor.))]
             (->> (xsssf-color-fill-by-css color background-color)
                (.setFillBackgroundColor cs)))))))
  (when-let [cs-f (and font
                       (.getFont cs))]
    (fill-font cs-f font)
    (.setFont cs cs-f))
  (some->> fmt
           ((fn [fmt]
             (if (:s fmt)
               (println "DataFormat with string is not supported this version")
               (.setDataFormat cs ^Short (:n fmt))))))
  (some->> wrap-text
           (.setWrapText cs))
  (some->> horizontal-alignment
           keyword->horizontal-alignment
           (.setAlignment cs))
  (some->> vertical-alignment
           keyword->vertical-alignment
           (.setVerticalAlignment cs)))


(defmulti set-cell-value (fn [^XSSFCell c {:keys [f v]}]
                           (cond
                             f CellType/FORMULA
                             (string? v) CellType/STRING
                             (boolean? v) CellType/BOOLEAN
                             (number? v) CellType/NUMERIC
                             (= Date
                                (class v)) CellType/NUMERIC
                             :else CellType/BLANK)))
(defmethod set-cell-value :default
  [_ _])
(defmethod set-cell-value CellType/FORMULA
  [c {:keys [f]}]
  (.setCellFormula c f))
(defmethod set-cell-value CellType/STRING
  [c {:keys [v]}]
  (.setCellValue c ^String v))
(defmethod set-cell-value CellType/BOOLEAN
  [c {:keys [v]}]
  (.setCellValue c ^Boolean v))
(defmethod set-cell-value CellType/NUMERIC
  [c {:keys [v]}]
  (cond (= Date
           (class v))
        (.setCellValue c ^Date v)

        (DateUtil/isCellDateFormatted c)
        (->> v
             (Date.)
             (.setCellValue c))

        :else
        (.setCellValue c (double v))))


(defn fill-cell
  [^Workbook wb ^Cell cell {:keys [v s f] :as cell-data}]
  (let [cell-s (.getCellStyle cell)]
    (set-cell-value cell cell-data)
    (when s
      (fill-style wb cell-s s)
      (.setCellStyle cell cell-s))))


(defn add-sheet! [^XSSFWorkbook wb sheet-data]
  (let [^XSSFSheet sheet (.createSheet wb)
        cells            (:rows sheet-data)
        start-row        (:start-row sheet-data 0)]
    (.setSheetName wb 0 (:name sheet-data))
    (doseq [row-i (range start-row (count cells))]
      (let [^XSSFRow row    (.createRow sheet (int row-i))
            row-data        (vec (get cells (- row-i start-row)))
            get-cell-height #(get-in % [:s :font :size] 11)]
        (doseq [col-i (range (count row-data))]
          (when (get row-data col-i)
            (let [^XSSFCell cell (.createCell row (int col-i))
                  cell-s         (.createCellStyle wb)
                  font           (.createFont wb)]
              (.setFont cell-s font)
              (.setCellStyle cell cell-s)
              (->> col-i
                   (get row-data)
                   (fill-cell wb cell)))))
        (->> row-data
             (apply max-key get-cell-height)
             get-cell-height
             (+ 6)
             (.setHeightInPoints row))))
    (->> sheet-data
         :col-styles
         (map-indexed (fn [indx style]
                        (let [{:keys [width]} style]
                          (cond
                            (int? width) (.setColumnWidth sheet indx (* 256 width))
                            (= width :auto) (.autoSizeColumn sheet indx false)))))
         doall)))


(defn ->excel-byte-array
  [sheets]
  (let [wb         (XSSFWorkbook.)
        _          (doseq [sheet-data sheets]
                     (add-sheet! wb sheet-data))
        _          (XSSFFormulaEvaluator/evaluateAllFormulaCells wb)
        out-stream (java.io.ByteArrayOutputStream.)
        _          (.write wb out-stream)
        ba         (.toByteArray out-stream)]
    (.close out-stream)
    ba))


(defn write-file [pathname byte-array]
  (let [^FileOutputStream fos (FileOutputStream. ^String pathname)]
    (.write fos (bytes byte-array))
    (.close fos)))


(do
  (def test-data
    (let [header-s  {:s {:font                 {:bold true
                                                :size 20}
                         :horizontal-alignment :center
                         :vertical-alignment   :center}}
          default-s {:s {:font {:size 12}}}]
      [{:name       "Pricing"
        :col-styles [{:width 20} {:width :auto}]
        :rows       (-> (->> ["Product" "Price"]
                             (mapv (partial assoc header-s :v)))
                        (cons (->> [["Windows" 1000]
                                    ["Norton Antivirus" 200]
                                    ["Microsoft Office" 100]]
                                   (mapv (fn [x]
                                           (mapv (partial assoc default-s :v) x)))))
                        (concat [(->> [{:v "Total"} {:f "SUM(B2:B4)"}]
                                      (mapv (partial merge header-s)))])
                        vec)}]))


  (defn test-export
    ([] (test-export "text.xlsx"))
    ([filepath]
     (->> test-data
          ->excel-byte-array
          (write-file filepath)))))

