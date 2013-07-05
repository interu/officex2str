# -*- coding: utf-8 -*-
require 'spec_helper'

describe Officex2str do
  context "#valid_file?" do
    subject do
      Officex2str.new(@file_path).send(:valid_file?)
    end
    context "extname is docx" do
      before { @file_path = "fixtures/sample.docx" }
      it { subject.should be_true }
    end
    context "extname is xlsx" do
      before { @file_path = "fixtures/sample.xlsx" }
      it { subject.should be_true }
    end
    context "extname is pptx" do
      before { @file_path = "fixtures/sample.pptx" }
      it { subject.should be_true }
    end
    context "extname is txt" do
      before { @file_path = "fixtures/sample.txt" }
      it { subject.should be_false }
    end
  end

  context "#pickup_pages" do
    subject do
      archives = Zip::Archive.open(@file_path) { |archive| archive.map(&:name) }
      Officex2str.new(@file_path).send(:pickup_pages, archives).sort
    end
    context "extname is docx" do
      before { @file_path = "fixtures/sample.docx" }
      it { subject.should == ["word/document.xml"] }
    end

    context "extname is xlsx" do
      before { @file_path = "fixtures/sample.xlsx" }
      it { subject.should == ["xl/comments1.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"] }
    end

    context "extname is pptx" do
      before { @file_path = "fixtures/sample.pptx" }
      it { subject.should == ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"] }
    end

  end

  context "#convert" do
    subject do
      Officex2str.convert(@file_path)
    end
    context "extname is xlsx" do
      before { @file_path = "fixtures/sample.xlsx" }
      it do
        subject.should include("複数シート対応")
        subject.should include("ソニックガーデン")
        subject.should include("ＳＯＮＩＣＧＡＲＤＥＮ")
        subject.should include("株式会社")
        subject.should include("コメント")
        subject.should include("STG001")
        subject.should include("STG003")
        subject.should_not include("sonicgarden")
        subject.should_not include("sheet")
      end
    end

    context "extname is docx" do
      before { @file_path = "fixtures/sample.docx" }
      it do
        subject.should include("複数ページ対応")
        subject.should include("ソニックガーデン")
        subject.should include("テキストボックス")
        subject.should_not include("sonicgarden")
        subject.should_not include("sheet")
      end
    end

    context "extname is pptx" do
      before { @file_path = "fixtures/sample.pptx" }
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

    context "extname is txt" do
      before { @file_path = "fixtures/sample.txt" }
      it do
        lambda {
          Officex2str.convert("fixtures/sample.txt")
        }.should raise_error(Officex2str::InvalidFileTypeError, "Not recognized file type")
      end
    end
  end
end
