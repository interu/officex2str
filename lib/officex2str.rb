require 'nokogiri'
require 'zipruby'
require 'mime/types'
#require "officex2str/version"

module Officex2str
  def self.convert(file_path)
    archives   = Zip::Archive.open(file_path) { |archive| archive.map(&:name) }
    pages      = self.pickup_pages(file_path, archives)
    xmls       = self.extract_xmls(file_path, pages)
    self.xml_to_str(xmls)
  end

private
  def self.pickup_pages file_path, archives
    case content_type = MIME::Types.type_for(file_path).first.content_type
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

  def self.extract_xmls file_path, pages
    xml_text = []
    Zip::Archive.open(file_path) { |archive| pages.each{ |page| archive.fopen(page) do |f| xml_text << f.read end; } }
    xml_text
  end

  def self.xml_to_str xml_text
    text = ""
    xml_text.each{|xml_t| text << Nokogiri.XML(xml_t.toutf8, nil, 'utf8').to_str } unless xml_text.empty?
    text
  end
end
