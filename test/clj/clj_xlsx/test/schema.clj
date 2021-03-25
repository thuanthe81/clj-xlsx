(ns clj-xlsx.test.schema
  (:require [clojure.test :refer [deftest is testing]]
            [clj-xlsx.schema :as sch]
            [clojure.spec.alpha :as s]
            [spec-tools.data-spec :as ds]
            [clojure.string :as str]))

(def date-ts 61538659200000) ;; "2020-01-01T00:00:00Z

(def test-data
  {:case-1 [{:name "Pricing"
             :rows [[{:v "Product"
                      :s {:font                 {:name "Arial"
                                                 :bold true
                                                 :size 20}
                          :horizontal-alignment :center
                          :vertical-alignment   :center
                          :background-color     "#D5DBD6"}}
                     nil
                     {:v "Price"
                      :s {:font                 {:name "Arial"
                                                 :bold true
                                                 :size 20}
                          :horizontal-alignment :center
                          :vertical-alignment   :center}}]]}]
   :case-2 [{:name       "Formula"
             :col-styles [nil {:width 30} {:width :auto}]
             :rows       [[{:v 15
                            :s {:font {:name "Arial"
                                       :size 12}}}
                           {:v 54
                            :s {:font {:name "Arial"
                                       :size 12}}}
                           {:v 69
                            :f "SUM(A1:A2)"
                            :s {:font {:name   "Arial"
                                       :size   16
                                       :color  "#FF0000"
                                       :italic true}}}]]}]
   :case-3 [{:name       "Formatting"
             :col-styles [{:width :auto}]
             :rows       [[{:v   date-ts
                            :s   {:font {:name "Arial"
                                         :size 12}
                                  :fmt {:n 0xe}}}]]}]})

(deftest schema-test
  (->> test-data
       (mapv (fn [[case-id data]]
               (let [label (->> case-id
                                name
                                ((fn [s] (str/split s #"-")))
                                (apply format "Testing schema %s %s"))]
                 (testing label
                   (is (s/conform (ds/spec ::t sch/sheets) data))))))))