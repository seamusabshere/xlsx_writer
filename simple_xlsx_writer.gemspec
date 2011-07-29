# -*- encoding: utf-8 -*-
$LOAD_PATH.unshift 'lib'

Gem::Specification.new do |s|
  s.name        = "mumboe-simple_xlsx_writer"
  s.version     = '0.5.10'
  s.platform    = Gem::Platform::RUBY
  s.authors     = ["Dee Zsombor", "Justin Beck", "Michael Alletto"]
  s.email       = ["dbortz@gmail.com", "mikea@mumboe.com"]
  s.homepage    = "http://github.com/dbortz/simple_xlsx_writer"
  s.summary     = "Gem version of Justin Beck's modifications to Dee Zsombor's Simple XLSX writer to use Tempfile"
  s.description = "Writes XLSX files"
 
  s.files        = Dir.glob("{bin,lib,test}/**/*") + %w(LICENSE README Rakefile)
  s.add_dependency("fast_xs", ">= 0.7.3")
  s.add_dependency("zip", ">= 2.0.2")
end
