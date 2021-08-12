require 'spec_helper'

describe ToSpreadsheet::Rule::Format do
  let :spreadsheet do
    build_spreadsheet haml: <<-HAML
:ruby
  format_xls do
    format column: 0, width: 25
    format 'tr', fg_color: lambda { |row| 'cccccc' if row.row_index.odd? }
    format 'table' do |ws|
      ws.name = 'Test'
    end
  end
%table
  %tr
    %th
  %tr
    %td
    HAML
  end

  let(:sheet) { spreadsheet.workbook.worksheets[0] }

  context 'local styles' do
    it 'sets column width' do
      expect(sheet.column_info[0].width).to eq(25)
    end
    it 'runs lambdas on properties' do
      cell = sheet.rows[1].cells[0]
      styles = sheet.workbook.styles
      font_id = styles.cellXfs[cell.style].fontId
      expect(styles.fonts[font_id].color.rgb).to eq(Axlsx::Color.new(rgb: 'cccccc').rgb)
    end
    it 'runs blocks for selectors' do
      expect(sheet.name).to eq('Test')
    end
  end
end
