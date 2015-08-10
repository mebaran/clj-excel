(defproject com.outpace/clj-excel "0.0.8"
  :description "Excel bindings for Clojure, based on Apache POI."
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.apache.poi/poi "3.9"]
                 [org.apache.poi/poi-ooxml "3.9"]]
  :profiles {:dev {:source-paths ["dev"]
                   :resource-paths ["test-resources"]
                   :plugins [[lein-cloverage "1.0.2"]]
                   :dependencies [[org.clojure/tools.namespace "0.2.4"]]
                   :global-vars {*print-length* 20}}
             :test {:resource-paths ["test-resources"]}}
  :global-vars {*warn-on-reflection* true}
  :deploy-repositories [["clojars" {:url "https://clojars.org/repo"
                                    :sign-releases false}]]
  :release-tasks [["vcs" "assert-committed"]
                  ["change" "version" "leiningen.release/bump-version" "release"]
                  ["vcs" "commit"]
                  ["deploy"]
                  ["change" "version" "leiningen.release/bump-version"]
                  ["vcs" "commit"]])
