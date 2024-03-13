# frozen_string_literal: true
require 'spec_helper'

describe Workbook do

  let(:filename) { "#{File.dirname(__FILE__)}/data/Spec.xlsx" }

  it 'Reads from StringIO' do
    io = StringIO.new File.read(filename)
    Workbook.open io do |w|
      w.should have(5).sheets
    end
  end

  it 'Sheets count' do
    Workbook.open filename do |w|
      w.should have(5).sheets
    end
  end

  it 'Sheet names' do
    Workbook.open filename do |w|
      w.sheet_names.should eq %w(test_otros test_spec test_param Lenguajes ont_demo)
    end
  end

  it 'Find sheet by index' do
    Workbook.open filename do |w|
      w.sheets[0].name.should eq 'test_otros'
      w.sheets[1].name.should eq 'test_spec'
      w.sheets[2].name.should eq 'test_param'
      w.sheets[3].name.should eq 'Lenguajes'
      w.sheets[4].name.should eq 'ont_demo'
    end
  end

  it 'Find sheet by name' do
    Workbook.open filename do |w|
      w.sheets('test_otros').name.should eq 'test_otros'
      w.sheets('test_spec').name.should eq 'test_spec'
      w.sheets('test_param').name.should eq 'test_param'
      w.sheets('Lenguajes').name.should eq 'Lenguajes'
      w.sheets('ont_demo').name.should eq 'ont_demo'
    end
  end

  it 'Shared strings' do
    Workbook.open filename do |w|
      w.should have(56).shared_strings
      w.shared_strings.should include 'LevenshteinDistance'
      w.shared_strings.should include 'TST_ModMan_Insulto_SU_Normal'
    end
  end
end
