require 'spec_helper'

describe ToSpreadsheet::Renderer do
  let :spreadsheet do
    build_spreadsheet haml: <<-HAML
%table
%table
    HAML
  end

  context 'worksheets' do
    it 'are created 1 per <table>' do
      expect(spreadsheet.workbook.worksheets.length).to eq(2)
    end
  end
end
