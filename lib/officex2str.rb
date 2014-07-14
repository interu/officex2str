require 'nokogiri'
require 'mime/types'
require 'zip'

class Officex2str
  class InvalidFileTypeError < Exception; end

  DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  VALID_CONTENT_TYPE = [DOCX_CONTENT_TYPE, XLSX_CONTENT_TYPE, PPTX_CONTENT_TYPE].freeze

  attr_accessor :path, :content_type, :entries, :xmls

  def self.convert(file_path)
    self.new(file_path).convert
  end

  def initialize(file_path)
    @path = file_path
    @content_type = MIME::Types.type_for(path).first.content_type
    @entries = valid_file? ? Zip::File.open(path).entries : []
    @xmls = []
  end

  def convert
    if valid_file?
      extract_xmls
      xml_to_str
    else
      raise InvalidFileTypeError, "Not recognized file type"
    end
  end

private
  def valid_file?
    !!VALID_CONTENT_TYPE.include?(content_type)
  end

  def select_target_entries
    case content_type
    when DOCX_CONTENT_TYPE
      entries.select{|a| /^word\/document/ =~ a.to_s}
    when XLSX_CONTENT_TYPE
      entries.select{|a| /^xl\/worksheets\/sheet/ =~ a.to_s or /^xl\/sharedStrings/ =~ a.to_s or /^xl\/comments/ =~ a.to_s }
    when PPTX_CONTENT_TYPE
      entries.select{|a| /^ppt\/slides\/slide/ =~ a.to_s}
    else
      raise InvalidContentTypeError, "Not recognized content type"
    end
  end

  def extract_xmls
    select_target_entries.map{|entry| xmls << entry.get_input_stream.read }
  end

  def xml_to_str
    return '' if xmls.empty?
    xmls.inject(""){|result, xml| result << Nokogiri.XML(xml, nil, 'utf8').to_str }
  end
end
