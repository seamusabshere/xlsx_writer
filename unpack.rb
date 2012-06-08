#!/usr/bin/env ruby

# Usage: ./unpack.rb foo.xlsx
# outputs to foo/

# requires xmllint to be in your path

require 'unix_utils'

src = File.expand_path ARGV[0]
dest = File.join(File.dirname(src), File.basename(src, '.xlsx').gsub(/\W/, '_'))

raise "#{dest} exists" if File.exist?(dest)
tmp_path = UnixUtils.unzip src
FileUtils.mv tmp_path, dest

Dir["#{dest}/**/*"].each do |infile|
  if File.file?(infile) and not File.extname(infile) == '.vml' and File.read(infile, 50).include?('<?xml')
    outfile = infile + '.tmp'
    `xmllint --format #{infile} > #{outfile}`
    FileUtils.mv outfile, infile
  end
end
