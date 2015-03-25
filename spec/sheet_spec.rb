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
        s.rows[0].should eq ['LevenshteinDistance', 0]
        s.rows[1].should eq ['Case sensitive', false]
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
      csv[0].should eq "\"LevenshteinDistance\",\"0.0\"\n"
      csv[1].should eq "\"Case sensitive\",\"false\"\n"
      csv[2].should eq "\"Fields\",\"Type\",\"URL Mining\"\n"
      csv[3].should eq "\"autor\",\"text\",\"false\"\n"
      csv[4].should eq "\"texto\",\"text\",\"false\"\n"
      csv[5].should eq "\"url\",\"text\",\"false\"\n"
      csv[6].should eq "\"comentario\",\"text\",\"false\"\n"
    end
  end

end