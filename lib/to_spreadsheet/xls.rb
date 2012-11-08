module ToSpreadsheet
  require 'spreadsheet'
  module XLS
    extend self

    def to_io(html)
      spreadsheet = Spreadsheet::Workbook.new
      Nokogiri::HTML::Document.parse(html).css('table').each_with_index do |xml_table, i|
        sheet = spreadsheet.create_worksheet(:name => xml_table.css('caption').inner_text.presence || "Sheet #{i + 1}")
        xml_table.css('tr').each_with_index do |row_node, row|
          row_node.css('[width]').each_with_index do |col_node, col|
            set_column_width sheet.column(col), col_node[:width]
          end
          row_node.css('th,td').each_with_index do |col_node, col|
            sheet[row, col] = typed_node_val(col_node)
          end
        end
      end
      io = StringIO.new
      spreadsheet.write(io)
      io.rewind
      io
    end

  private

    def typed_node_val(node)
      val = val_or_default(node)
      begin
        case node[:class]
          when /decimal|float/
            val.to_f
          when /num|int/
            val.to_i
          when /datetime/
            DateTime.parse(val)
          when /date/
            Date.parse(val)
          when /time/
            Time.parse(val)
          else
            val
        end
      rescue
        val
      end
    end

    def val_or_default(node)
      val = node.inner_text
      if val.blank?
        node['data-default']
      else
        val
      end
    end

    def set_column_width(column, width)
      column.width = width
    end

  end
end
