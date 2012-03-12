require 'spec_helper'

describe ToSpreadsheet::XLS do
  let(:spreadsheet) {
    html  = Haml::Engine.new(TEST_HAML).render
    xls_io = ToSpreadsheet::XLS.to_io(html)
    Spreadsheet.open(xls_io)
  }

  it 'creates multiple worksheets' do
    spreadsheet.should have(2).worksheets
  end

  it 'supports num format' do
    spreadsheet.worksheet(0)[1, 1].should == 20
  end

  it 'supports date format' do
    spreadsheet.worksheet(0)[1, 2].class.should be(Date)
  end

  # This is for final manual test
  # The test spreadsheet will be saved to /tmp/spreadsheet.xls
  it 'writes to disk' do
    f = File.open('/tmp/spreadsheet.xls', 'wb')
    Spreadsheet.writer(f).write_workbook(spreadsheet, f)
    f.close
  end

end

TEST_HAML = <<-HAML

%table
  %caption A worksheet
  %thead
    %tr
      %th Name
      %th Age
      %th Date
  %tbody
    %tr
      %td Gleb
      %td.num 20
      %td.date 27/05/1991
    %tr
      %td John
      %td.num 21
      %td.date 01/05/1990

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
      %td.num 19
      %td.date 10/05/1991

HAML