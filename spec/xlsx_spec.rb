require 'spec_helper'

describe ToSpreadsheet::XLSX do
  let(:spreadsheet) {
    html  = Haml::Engine.new(TEST_HAML).render
    package = ToSpreadsheet::XLSX.generate(html)
  }

  it 'creates multiple worksheets' do
    spreadsheet.workbook.should have(2).worksheets
  end

  it 'supports num format' do
    spreadsheet.workbook.worksheets[0].rows[1].cells[1].value.should == 20
  end

  it 'support float format' do
    spreadsheet.workbook.worksheets[1].rows[1].cells[1].type.should be(:float)
  end

  it 'supports date format' do
    spreadsheet.workbook.worksheets[0].rows[1].cells[2].type.should be(:date)
  end

  it 'parses null dates' do
    spreadsheet.workbook.worksheets[0].rows[2].cells[2].type.should_not be(:date)
  end

  it 'parses default values' do
    spreadsheet.workbook.worksheets[0].rows[2].cells[1].value.should == 100
  end

  it 'sets column width based on th width' do
    spreadsheet.workbook.worksheets[0].column_info[0].width.should == 25
  end

  it 'sets column width based on td width' do
    spreadsheet.workbook.worksheets[1].column_info[1].width.should == 35
  end

  # The test spreadsheet will be saved to /tmp/spreadsheet.xls
  it 'writes to disk' do
    spreadsheet.serialize('/tmp/spreadsheet.xlsx')
    File.exists?('/tmp/spreadsheet.xlsx').should == true
  end

end

TEST_HAML = <<-HAML

%table
  %caption A worksheet
  %thead
    %tr
      %th{ width: 25 } Name
      %th Age
      %th Date
  %tbody
    %tr
      %td Gleb
      %td.num 20
      %td.date 27/05/1991
    %tr
      %td John
      %td.num{ data: { default: 100 } }
      %td.date

%table
  %caption Another worksheet
  %thead
    %tr
      %th Name
      %th Age
      %th Date
  %tbody
    %tr
      %td Alice
      %td.float{ width: 35 } 19.5
      %td.date 10/05/1991

HAML