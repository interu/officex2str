require 'nokogiri'
require 'zipruby'
require 'mime/types'

class Officex2str
  class InvalidFileTypeError < Exception; end

  DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  VALID_CONTENT_TYPE = [DOCX_CONTENT_TYPE, XLSX_CONTENT_TYPE, PPTX_CONTENT_TYPE].freeze

  attr_accessor :path, :content_type

  def self.convert(file_path)
    self.new(file_path).convert
  end

  def initialize(file_path)
    @path = file_path
    @content_type = MIME::Types.type_for(path).first.content_type
  end

  def convert
    if valid_file?
      archives   = Zip::Archive.open(path) { |archive| archive.map(&:name) }
      pages      = pickup_pages(archives)
      xmls       = extract_xmls(pages)
      xml_to_str(xmls)
    else
      raise InvalidFileTypeError, "Not recognized file type"
    end
  end

private
  def valid_file?
    !!VALID_CONTENT_TYPE.include?(content_type)
  end

  def pickup_pages archives
    case content_type
    when DOCX_CONTENT_TYPE
      archives.select{|a| /^word\/document/ =~ a}
    when XLSX_CONTENT_TYPE
      archives.select{|a| /^xl\/worksheets\/sheet/ =~ a or /^xl\/sharedStrings/ =~ a or /^xl\/comments/ =~ a }
    when PPTX_CONTENT_TYPE
      archives.select{|a| /^ppt\/slides\/slide/ =~ a}
    else
      raise InvalidContentTypeError, "Not recognized content type"
    end
  end

  def extract_xmls pages
    xml_text = []
    Zip::Archive.open(path) { |archive| pages.each{ |page| archive.fopen(page) { |f| xml_text << f.read } } }
    xml_text
  end

  def xml_to_str xml_text
    text = ""
    xml_text.each{|xml_t| text << Nokogiri.XML(xml_t.toutf8, nil, 'utf8').to_str } unless xml_text.empty?
    text
  end
end
