require 'fileutils'
require 'tmpdir'
require 'posix/spawn'

module XlsxWriter
  module Utils
    def self.tmp_path(basename = nil, extname = nil)
      ::Kernel.srand
      ::File.join ::Dir.tmpdir, "XlsxWriter-#{basename}#{::Kernel.rand(99999999)}#{extname ? ".#{extname}" : ''}"
    end
    
    # zip -r -q #{filename} .
    def self.zip(src_dir)
      out_path = tmp_path('zip', 'zip')
      child = ::POSIX::Spawn::Child.new 'zip', '--recurse-paths', out_path, '.', :chdir => src_dir
      if child.success?
        out_path
      else
        raise ::RuntimeError, child.err
      end
    end

    # use awk to convert CR?LF to CRLF
    def self.unix2dos(path)
      out_path = tmp_path
      ::File.open(out_path, 'w') do |out|
        pid = ::POSIX::Spawn.spawn 'awk', '{ sub(/\r?$/,"\r"); print }', path, :out => out
        ::Process.waitpid pid
      end
      ::FileUtils.mv out_path, path
      path
    end
  end
end
