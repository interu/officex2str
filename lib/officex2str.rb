require 'nokogiri'
require 'zipruby'
require 'mime/types'

class Officex2str
  attr_accessor :path

  def self.convert(file_path)
    self.new(file_path).convert
  end

  def initialize(file_path)
    @path = file_path
  end

  def convert
     archives   = Zip::Archive.open(path) { |archive| archive.map(&:name) }
     pages      = pickup_pages(archives)
     xmls       = extract_xmls(pages)
     xml_to_str(xmls)
  end

private
  def pickup_pages archives
    case content_type = MIME::Types.type_for(path).first.content_type
    when "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      archives.select{|a| /^word\/document/ =~ a}
    when "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      archives.select{|a| /^xl\/worksheets\/sheet/ =~ a or /^xl\/sharedStrings/ =~ a or /^xl\/comments/ =~ a }
    when "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      archives.select{|a| /^ppt\/slides\/slide/ =~ a}
    else
      nil
    end
  end

  def extract_xmls pages
    xml_text = []
    Zip::Archive.open(path) { |archive| pages.each{ |page| archive.fopen(page) do |f| xml_text << f.read end; } }
    xml_text
  end

  def xml_to_str xml_text
    text = ""
    xml_text.each{|xml_t| text << Nokogiri.XML(xml_t.toutf8, nil, 'utf8').to_str } unless xml_text.empty?
    text
  end
end
