(ns clj-xlsx.share.schema
  (:require [schema.core :as s]))

(def css-color (s/pred (fn [s]
                           (->> s
                                (re-find #"(^#[0-9a-fA-F]{8}|^#[0-9a-fA-F]{6})$")
                                first
                                some?))))

(def font
  {(s/optional-key :color)  css-color
   (s/optional-key :name)   s/Str
   (s/optional-key :size)   s/Num
   (s/optional-key :bold)   s/Bool
   (s/optional-key :italic) s/Bool})

(def data-format-num     ;; Refer number from org.apache.poi.ss.usermodel.BuiltinFormats
  {(s/optional-key :n) s/Num})

(def data-format-str
  {:s                      s/Str
   (s/optional-key :local) s/Str})

(def cell-style
  {(s/optional-key :font)                 font
   (s/optional-key :background-color)     css-color
   (s/optional-key :fmt)                  (s/cond-pre data-format-num
                                                      data-format-str)
   (s/optional-key :horizontal-alignment) s/Keyword
   (s/optional-key :vertical-alignment)   s/Keyword})

(def cell
  {(s/optional-key :v) s/Any
   (s/optional-key :f) s/Str
   (s/optional-key :s) cell-style})

(def col-style
  (s/if map?
    {:width (s/cond-pre s/Num
                        (s/eq :auto))}
    (s/eq nil)))

(def row
  [(s/maybe cell)])

(def rows
  [row])

(def sheet
  {:name                       s/Str
   (s/optional-key :start-row) s/Num
   (s/optional-key :col-styles)                 [col-style]
   :rows                       rows})

(def sheets
  [sheet])