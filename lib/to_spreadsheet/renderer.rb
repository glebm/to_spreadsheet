require 'axlsx'
require 'nokogiri'

module ToSpreadsheet
  module Renderer
    extend self

    def to_stream(html, context = nil)
      to_package(html, context).to_stream
    end

    def to_data(html, context = nil)
      to_package(html, context).to_stream.read
    end

    def to_package(html, context = nil)
      context ||= ToSpreadsheet::Context.global.merge(Context.new)
      package = build_package(html, context)
      context.rules.each do |rule|
        #Rails.logger.debug "Applying #{rule}"
        rule.apply(context, package)
      end
      package
    end

    private

    def build_package(html, context)
      package     = ::Axlsx::Package.new
      spreadsheet = package.workbook
      doc         = Nokogiri::HTML::Document.parse(html)
      # Workbook <-> %document association
      context.assoc! spreadsheet, doc
      doc.css('table').each_with_index do |xml_table, i|
        sheet = spreadsheet.add_worksheet(
            name: xml_table.css('caption').inner_text.presence || xml_table['name'] || "Sheet #{i + 1}"
        )
        # Sheet <-> %table association
        context.assoc! sheet, xml_table
        xml_table.css('tr').each do |row_node|
          xls_row = sheet.add_row
          # Row <-> %tr association
          context.assoc! xls_row, row_node
          row_node.css('th,td').each do |cell_node|
            xls_col = xls_row.add_cell cell_node.inner_text
            # Cell <-> th or td association
            context.assoc! xls_col, cell_node
          end
        end
      end
      package
    end
  end
end
