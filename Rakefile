require 'rubygems'
require 'rake'
require 'rake/testtask'
require 'rake/gempackagetask'

task :default => [:test]

Rake::TestTask.new do |test|
  test.libs       << "test"
  test.test_files =  Dir['test/**/*_test.rb'].sort
  test.verbose    =  true
end

Rake::GemPackageTask.new(SPEC) do |pkg|
  pkg.need_tar = true
end
