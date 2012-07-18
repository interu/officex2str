# -*- coding: utf-8 -*-
require 'spec_helper'

describe Officex2str do
  context "#pickup_pages" do
    subject do
      archives = Zip::Archive.open(@file_path) { |archive| archive.map(&:name) }
      Officex2str.send(:pickup_pages, File.extname(@file_path), archives).sort
    end
    context "extname is docx" do
      before do
        @file_path = "fixtures/sample.docx"
      end
      it { subject.should == ["word/document.xml"] }
    end

    context "extname is xlsx" do
      before do
        @file_path = "fixtures/sample.xlsx"
      end
      it { subject.should == ["xl/comments1.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"] }
    end

    context "extname is pptx" do
      before do
        @file_path = "fixtures/sample.pptx"
      end
      it { subject.should == ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"] }
    end
  end

  context "#convert" do
    subject do
      archives = Zip::Archive.open(@file_path) { |archive| archive.map(&:name) }
      #pages = Officex2str.pickup_pages(File.extname(@file_path), archives)
      pages = Officex2str.send(:pickup_pages, File.extname(@file_path), archives)
      xmls = Officex2str.send(:extract_xmls, @file_path, pages)
      Officex2str.convert(@file_path)
    end
    context "extname is xlsx" do
      before do
        @file_path = "fixtures/sample.xlsx"
      end
      it do
        subject.should include("複数シート対応")
        subject.should include("ソニックガーデン")
        subject.should include("ＳＯＮＩＣＧＡＲＤＥＮ")
        subject.should include("株式会社")
        subject.should include("コメント")
        subject.should_not include("sonicgarden")
        subject.should_not include("sheet")
      end
    end

    context "extname is docx" do
      before do
        @file_path = "fixtures/sample.docx"
      end
      it do
        subject.should include("複数ページ対応")
        subject.should include("ソニックガーデン")
        subject.should include("テキストボックス")
        subject.should_not include("sonicgarden")
        subject.should_not include("sheet")
      end
    end

    context "extname is pptx" do
      before do
        @file_path = "fixtures/sample.pptx"
      end
      it do
        subject.should include("Aタイトル")
        subject.should include("Aサブタイトル")
        subject.should include("タイトルB")
        subject.should include("テキストB")
        subject.should include("テキストボックスB")
        subject.should_not include("sonicgarden")
        subject.should_not include("sheet")
      end
    end

  end
end