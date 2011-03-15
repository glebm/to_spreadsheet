module ToSpreadsheet
  require 'spreadsheet'
  module XLS
    extend self

    def to_io(html)
      spreadsheet = Spreadsheet::Workbook.new
      Nokogiri::XML::Document.parse(html).xpath('/table').each_with_index do |xml_table, i|
        sheet = spreadsheet.create_worksheet(:name => xml_table.css('caption').inner_text.presence || "Sheet #{i + 1}")
        xml_table.css('tr').each_with_index do |row_node, row|
          row_node.css('th,td').each_with_index do |col_node, col|
            text = col_node.inner_text
            sheet[row,col] = text
          end
        end
      end
      io = StringIO.new
      spreadsheet.write(io)
      io.rewind
      io
    end
  end
end