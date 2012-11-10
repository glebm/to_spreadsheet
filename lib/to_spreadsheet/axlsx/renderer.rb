require 'axlsx'
require 'to_spreadsheet/axlsx/formatter'
module ToSpreadsheet
  module Axlsx
    module Renderer
      include Formatter
      extend self

      def to_stream(html, context = ToSpreadsheet.context)
        to_package(html, context).to_stream
      end

      def to_data(html, context = ToSpreadsheet.context)
        to_package(html, context).to_stream.read
      end

      def to_package(html, context = ToSpreadsheet.context)
        package = build_package(html, context)
        apply_formats(package, context)
        # Don't leak memory: clear all dom <-> axslsx associations
        context.clear_assoc!
        # Numbers compat
        package.use_shared_strings = true
        package
      end

      private

      def build_package(html, context)
        package     = ::Axlsx::Package.new
        spreadsheet = package.workbook
        doc         = Nokogiri::HTML::Document.parse(html)
        context.assoc! spreadsheet, doc
        doc.css('table').each_with_index do |xml_table, i|
          sheet = spreadsheet.add_worksheet(
              name: xml_table.css('caption').inner_text.presence || xml_table['name'] || "Sheet #{i + 1}"
          )
          context.assoc! sheet, xml_table
          xml_table.css('tr').each do |row_node|
            xls_row = sheet.add_row
            context.assoc! xls_row, row_node
            row_node.css('th,td').each do |cell_node|
              xls_col = xls_row.add_cell  cell_node.inner_text
              context.assoc! xls_col, cell_node
            end
          end
        end
        package
      end
    end
  end
end
