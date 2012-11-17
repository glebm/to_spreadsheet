require 'spec_helper'



describe ToSpreadsheet::Rule::DefaultValue do
  let :spreadsheet do
    build_spreadsheet(haml: <<HAML)
- format_xls 'table' do
  - default 'td.price', 100
%table
  %tr
    %td.price
    %td.price 50
HAML
  end

  let :row do
    spreadsheet.workbook.worksheets[0].rows[0]
  end

  context 'default values' do
    it 'get set when the cell is empty' do
      row.cells[0].value.should == 100
    end
    it 'does not get set when the cell is not empty' do
      row.cells[1].value.should == 50
    end
  end
end