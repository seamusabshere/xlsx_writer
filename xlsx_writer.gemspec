# -*- encoding: utf-8 -*-
$LOAD_PATH.unshift 'lib'

Gem::Specification.new do |s|
  s.name        = "xlsx_writer"
  s.version     = '0.6.0'
  s.authors     = ["Dee Zsombor", "Justin Beck", "Seamus Abshere"]
  s.email       = ["seamus@abshere.net"]
  s.homepage    = "http://github.com/seamusabshere/xlsx_writer"
  s.summary     = "Refactored version of Justin Beck's modifications to Dee Zsombor's Simple XLSX writer"
  s.description = "Writes XLSX files"
  s.files        = Dir.glob("{bin,lib,test}/**/*") + %w(LICENSE README.markdown Rakefile)
  s.add_dependency 'activesupport'
  s.add_dependency "fast_xs", ">= 0.7.3"
  s.add_dependency "zip", ">= 2.0.2"
  s.add_dependency 'posix-spawn'
end
