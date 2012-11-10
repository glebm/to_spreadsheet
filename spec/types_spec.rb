require 'spec_helper'

describe ToSpreadsheet::Themes::Default do
  let :spreadsheet do
    build_spreadsheet haml: <<-HAML

%table
  %tr
    %td.num 20
    %td.float 1
    %td.date 27/05/1991
    %td.date
HAML
  end

  let :row do
    spreadsheet.workbook.worksheets[0].rows[0]
  end

  context 'data types' do
    it 'num' do
      row.cells[0].value.should == 20
    end

    it 'float' do
      row.cells[1].type.should be :float
    end

    it 'date' do
      row.cells[2].type.should be :date
    end

    it 'empty date' do
      row.cells[3].type.should_not be :date
    end
  end
end
