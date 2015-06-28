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
      expect(row.cells[0].value).to eq(20)
    end

    it 'float' do
      expect(row.cells[1].type).to be :float
    end

    it 'date' do
      expect(row.cells[2].type).to be :date
    end

    it 'empty date' do
      expect(row.cells[3].type).not_to be :date
    end
  end
end
