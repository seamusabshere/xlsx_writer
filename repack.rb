#!/usr/bin/env ruby

# Usage: ./repack.rb foo
# where "foo" is a directory previously unpacked to by unpack.rb
# outputs to out.xlsx

require 'unix_utils'

src = File.expand_path ARGV[0]
dest = File.join(File.dirname(src), 'out.xlsx')

raise "#{dest} exists" if File.exist?(dest)

src_copy = UnixUtils.tmp_path src
FileUtils.cp_r src, src_copy
Dir["#{src_copy}/**/*"].each do |infile|
  if File.file?(infile) and not File.extname(infile) == '.vml' and File.read(infile, 50).include?('<?xml')
    raise "uhh ohh #{File.dirname(infile)}" unless File.dirname(infile).start_with?(Dir.tmpdir)
    tmp_path = UnixUtils.unix2dos infile
    FileUtils.mv tmp_path, infile
  end
end

tmp_path = UnixUtils.zip src_copy
FileUtils.mv tmp_path, dest
