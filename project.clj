(defproject clj-excel "0.0.1"
  :description "Excel bindings for Clojure, based on Apache POI."
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.apache.poi/poi "3.9"]
                 [org.apache.poi/poi-ooxml "3.9"]]

  ;; lein with-profile dev cloverage [cloverage-opts]
  :profiles {:dev {:plugins [[lein-cloverage "1.0.2"]]}}
  :warn-on-reflection true)
