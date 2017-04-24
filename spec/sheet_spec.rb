# encoding: UTF-8
require 'spec_helper'

describe Sheet do

  let(:filename) { "#{File.dirname(__FILE__)}/data/Spec.xlsx" }
  let(:tmp_path) { "#{File.dirname(__FILE__)}/../tmp" }

  before :each do
    FileUtils.rm_rf tmp_path if Dir.exists? tmp_path
  end

  it 'Rows count' do
    Workbook.open filename do |w|
      w.sheets[0].rows.count.should eq 7
      w.sheets[1].rows.count.should eq 9
      w.sheets[2].rows.count.should eq 3
      w.sheets[3].rows.count.should eq 2
      w.sheets[4].rows.count.should eq 3
    end
  end

  it 'Rows collection' do
    Workbook.open filename do |w|
      w.sheets[0].rows.should be_an_instance_of RowsCollection
    end
  end

  it 'Rows content' do
    Workbook.open filename do |w|
      w.sheets[0].tap do |s|
        s.rows[0].should eq [
          'LevenshteinDistance',
          3.14,
          3,
          DateTime.new(2013, 12, 13, 8, 0, 58),
          DateTime.new(1970, 1, 1),
          BigDecimal.new('3.4028236692093801E+38'),
          DateTime.new(2015, 2, 13, 12, 40, 5)
        ]
        s.rows[1].should eq [
          'Case sensitive',
          false,
          3.0,
          DateTime.new(1970, 1, 1, 1, 0, 0)
        ]
        s.rows[2].should eq ['Fields', 'Type', 'URL Mining']
        s.rows[3].should eq ['autor', 'text', false]
        s.rows[4].should eq ['texto', 'text', false]
        s.rows[5].should eq ['url', 'text', false]
        s.rows[6].should eq ['comentario', 'text', false]
      end
    end
  end

  it 'Rows content skipping cells' do
    Workbook.open filename do |w|
      w.sheets[3].tap do |s|
        s.rows[0].should eq [nil, 'en', 'es', 'pt', 'un']
        s.rows[1].should eq ['default', 30, 50, 15, 5]
      end
    end
  end

  it 'Rows content with tag separators (>)' do
    Workbook.open filename do |w|
      w.sheets[4].tap do |s|
        s.rows[0].should eq ['Especificacion', 'Concepto/RegExp/Pair', 'ClienteTexto_Campos', 'ClienteTexto_Especificacion']
        s.rows[1].should eq ['Discriminación > Sexual | Insulto', 'puto', 'texto', 'TST_RechAuto_Insulto_SE_Normal']
        s.rows[2].should eq ['Insulto', 'boludo', 'texto', 'TST_ModMan_Insulto_SU_Normal']
      end
    end
  end

  it 'Export to CSV' do
    Workbook.open filename do |w|
      csv_file = "#{tmp_path}/#{w.sheets[0].name}.csv"

      File.should_not be_exists csv_file

      w.sheets[0].to_csv tmp_path

      csv = File.open(csv_file, 'r') { |f| f.readlines }
      # TODO: newer rubies use lowercase "e" in scientific numbers
      # csv[0].should eq %{"LevenshteinDistance","3.14","3","2013-12-13T08:00:58+00:00","1970-01-01T00:00:00+00:00","0.34028236692093801E39","2015-02-13T12:40:05+00:00"\n}
      csv[1].should eq %{"Case sensitive","false","3.0","1970-01-01T01:00:00+00:00"\n}
      csv[2].should eq "\"Fields\",\"Type\",\"URL Mining\"\n"
      csv[3].should eq "\"autor\",\"text\",\"false\"\n"
      csv[4].should eq "\"texto\",\"text\",\"false\"\n"
      csv[5].should eq "\"url\",\"text\",\"false\"\n"
      csv[6].should eq "\"comentario\",\"text\",\"false\"\n"
    end
  end

  context 'with 1904 date system' do
    let(:filename) { "#{File.dirname(__FILE__)}/data/Spec1904.xlsx" }

    it 'should use 1904 date system when converting dates' do
      Workbook.open filename do |w|
        w.sheets[0].tap do |s|
          s.rows[0].should eq [
            DateTime.new(1970, 1, 1, 1, 0, 0),
            DateTime.new(1970, 1, 1)
          ]
        end
      end
    end
  end

  context 'with inline strings' do
    let(:filename) { "#{File.dirname(__FILE__)}/data/SpecInlineStrings.xlsx" }

    it 'should read inline strings' do
      Workbook.open filename do |w|
        w.sheets[0].tap do |s|
          s.rows[0].should eq [
            'Test'
          ]
        end
      end
    end
  end

  context 'with number formats and auto format' do
    let(:filename) { "#{File.dirname(__FILE__)}/data/SpecNumberFormat.xlsx" }

    [ ["General",    "Test"],
      ["General",    123],
      ["General",    123.5],
      ["Fixnum",     123],
      ["Currency",   123.0],
      ["Date",       DateTime.new(1970, 1, 1)],
      ["Time",       DateTime.new(2015, 2, 13, 12, 40, 5)],
      ["Percentage", 0.9999],
      ["Fraction",   0.5],
      ["Scientific", BigDecimal.new('3.4028236692093801E+38')],
      ["Custom",     123.0],
    ].each.with_index do |row, i|
      name, value = row

      it "should typecast #{name}" do
        Workbook.open filename do |w|
          w.sheets[0].tap do |s|
            expect(s.rows[i+1]).to eq([name, value, "Test"])
          end
        end
      end
    end
  end

  context 'with number formats and without auto format' do
    let(:filename) { "#{File.dirname(__FILE__)}/data/SpecNumberFormat.xlsx" }

    [ ["General",    "Test"],
      ["General",    "0123"],
      ["General",    "0123.50"],
      ["Fixnum",     123],
      ["Currency",   123.0],
      ["Date",       DateTime.new(1970, 1, 1)],
      ["Time",       DateTime.new(2015, 2, 13, 12, 40, 5)],
      ["Percentage", 0.9999],
      ["Fraction",   0.5],
      ["Scientific", BigDecimal.new('3.4028236692093801E+38')],
      ["Custom",     123.0],
    ].each.with_index do |row, i|
      name, value = row

      it "should typecast #{name}" do
        Workbook.open filename, auto_format: false do |w|
          w.sheets[0].tap do |s|
            expect(s.rows[i+1]).to eq([name, value, "Test"])
          end
        end
      end
    end
  end
end
