(ns clj-xlsx.schema
  (:require [spec-tools.data-spec :as s]))

(defn css-color?
  [s]
  (->> s
       (re-find #"(^#[0-9a-fA-F]{8}|^#[0-9a-fA-F]{6})$")
       first
       some?))

(def font
  {(s/opt :color)  css-color?
   (s/opt :name)   string?
   (s/opt :size)   number?
   (s/opt :bold)   boolean?
   (s/opt :italic) boolean?})


(defn horizontal-alignment? [x] (some? (#{:left :right :center :justify} x)))
(defn vertical-alignment? [x] (some? (#{:top :bottom :center :justify} x)))


(def data-format
  {(s/opt :locale) string?
   (s/opt :f) string?                                       ;; not support this version
   (s/opt :n) number?}                                      ;; Refer number from org.apache.poi.ss.usermodel.BuiltinFormats
  )


(def cell-style
  {(s/opt :font) font
   (s/opt :background-color) css-color?
   (s/opt :fmt) data-format
   (s/opt :horizontal-alignment) horizontal-alignment?
   (s/opt :vertical-alignment) vertical-alignment?})


(def cell
  {(s/opt :v) any?
   (s/opt :f) string?
   (s/opt :s) cell-style})

(defn number-or-auto?
  [x]
  (or (#{:auto} x)
      (number? x)))
(def col-styles [(s/maybe {:width number-or-auto?})])


(def row
  [(s/maybe cell)])

(def rows
  [row])

;;(s/def ::name string?)
;;(s/def ::start-row ::num)
(def sheet
  {:name string?
   (s/opt :col-styles) col-styles
   (s/opt :start-row) number?
   :rows rows})

(def sheets
  [sheet])