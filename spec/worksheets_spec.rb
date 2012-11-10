require 'spec_helper'

describe ToSpreadsheet::Axlsx::Renderer do
  let :spreadsheet do
    build_spreadsheet haml: <<-HAML
%table
%table
    HAML
  end

  context 'worksheets' do
    it 'are created 1 per <table>' do
      spreadsheet.workbook.should have(2).worksheets
    end
  end
end
