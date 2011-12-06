# -*- encoding: utf-8 -*-
$LOAD_PATH.unshift 'lib'

Gem::Specification.new do |s|
  s.name        = "xlsx_writer"
  s.version     = '0.1.0'
  s.authors     = ["Dee Zsombor", "Justin Beck", "Seamus Abshere"]
  s.email       = ["seamus@abshere.net"]
  s.homepage    = "http://github.com/seamusabshere/xlsx_writer"
  s.summary     = "Writes XLSX files. Minimal XML and style. Supports autofilters and headers/footers with images and page numbers."
  s.description = "Writes XLSX files. Minimal XML and style. Supports autofilters and headers/footers with images and page numbers."
  s.files        = Dir.glob("{bin,lib,test}/**/*") + %w(LICENSE README.markdown Rakefile)
  s.add_dependency 'activesupport'
  s.add_dependency "fast_xs", ">= 0.7.3"
  s.add_dependency 'posix-spawn'
end
