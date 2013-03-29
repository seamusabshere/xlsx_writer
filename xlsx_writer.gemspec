# -*- encoding: utf-8 -*-
require File.expand_path('../lib/xlsx_writer/version', __FILE__)

Gem::Specification.new do |s|
  s.name        = "xlsx_writer"
  s.version     = XlsxWriter::VERSION
  s.authors     = ["Dee Zsombor", "Justin Beck", "Seamus Abshere"]
  s.email       = ["seamus@abshere.net"]
  s.homepage    = "https://github.com/seamusabshere/xlsx_writer"
  s.summary     = %{Writes XLSX files. Minimal XML and style. Supports autofilters and headers/footers with images and page numbers.}
  s.description = %{Writes XLSX files. Minimal XML and style. Supports autofilters and headers/footers with images and page numbers.}
  
  s.add_runtime_dependency 'activesupport'
  s.add_runtime_dependency 'fast_xs'
  s.add_runtime_dependency 'unix_utils'

  s.add_development_dependency 'minitest'
  s.add_development_dependency 'minitest-reporters'
  s.add_development_dependency 'yard'
  s.add_development_dependency 'remote_table'
  s.add_development_dependency 'ruby-decimal'
  
  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
end
