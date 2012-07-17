require 'digest/md5'

class XlsxWriter
  class SharedStrings
    BUFSIZE = 131072 #128kb

    attr_reader :document
    attr_reader :path
    attr_reader :indexes

    def initialize(document)
      @mutex = Mutex.new
      @document = document
      @indexes = {}
      @path = File.join document.staging_dir, relative_path
      FileUtils.mkdir_p File.dirname(path)
      @strings_tmp_file_writer = File.open(strings_tmp_file_path, 'wb')
    end

    def relative_path
      'xl/sharedstrings.xml'
    end

    def ndx(str)
      @mutex.synchronize do
        digest = Digest::MD5.digest str
        unless ndx = indexes[digest]
          ndx = indexes.length
          indexes[digest] = ndx
          @strings_tmp_file_writer.write %{<si><t>#{str.fast_xs}</t></si>}
        end
        ndx
      end
    end

    def generated?
      @generated == true
    end

    def generate
      return if generated?
      @mutex.synchronize do
        return if generated?
        @generated = true
        @strings_tmp_file_writer.close
        File.open(path, 'wb') do |f|
          f.write <<-EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{indexes.length}" uniqueCount="#{indexes.length}">
EOS
          File.open(strings_tmp_file_path, 'rb') do |strings_tmp_file_reader|
            buffer = ''
            while strings_tmp_file_reader.read(BUFSIZE, buffer)
              f.write buffer
            end
          end
          f.write %{</sst>}
        end
        File.unlink strings_tmp_file_path
      end
    end

    private

    def strings_tmp_file_path
      path + '.strings_tmp_file'
    end

  end
end
