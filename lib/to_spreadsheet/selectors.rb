module ToSpreadsheet
  # This is the DSL context for `format_xls`
  # Query types: :css, :column, :row or :range
  # Query values:
  # For css: [String] css selector
  # For column and row: [Fixnum] column/row number
  # For range: [String] table range, e.g. A4:B5
  module Selectors
    # Flexible API query match
    # Options (all optional):
    #  xls_worksheet
    #  xls_entity
    #  xml_document
    #  xml_node
    #  selector_type :css, :column, :row or :range
    #  selector_query
    def query_match?(options)
      return true if !options[:selector_query]
      case options[:selector_type]
        when :css
          css_match? options[:selector_query], options[:xml_document], options[:xml_node]
        when :column
          return false unless [Axlsx::Row, Axlsx::Cell].include?(options[:xml_node].class)
          column_number_match? options[:selector_query], options[:xml_node]
        when :row
          return false unless Axlsx::Cell == options[:xml_node].class
          row_number_match? options[:selector_query], options[:xml_node]
        when :range
          return false if entity.is_a?(Axlsx::Cell)
          range_contains? options[:selector_query], options[:xml_node]
        else
          raise "Unsupported type #{options[:selector_type].inspect} (:css, :column, :row or :range expected)"
      end
    end

    def each_cell(context, sheet, selector_type, selector_query, &block)
      if !selector_type
        sheet.rows.each do |row|
          sheet.cells.each do |cell|
            block.(cell)
          end
        end
        return
      end
      case selector_type
        when :css
          css_match selector_query, context.to_xml_node(sheet) do |xml_node|
            block.(context.to_xls_entity(xml_node))
          end
        when :column
          sheet.cols[selector_query].cells.each(&block)
        when :row
          sheet.cols[selector_query].cells.each(&block)
        when :range
          sheet[range].each(&block)
      end
    end

    def css_match(css_selector, xml_node, &block)
      xml_node.css(css_selector).each(&block)
    end

    def css_match?(css_selector, xml_document, xml_node)
      xml_document.css(css_selector).include?(xml_node)
    end

    def row_number_match?(row_number, xls_row_or_cell)
      if xls_row_or_cell.is_a? Axlsx::Row
        row_number == xls_row_or_cell.index
      elsif xls_row_or_cell.is_a? Axlsx::Cell
        row_number == xls_row_or_cell.row.index
      end
    end

    def column_number_match?(column_number, xls_cell)
      xls_cell.index == column_number if xls_cell.is_a?(Axlsx::Cell)
    end

    def range_match(range, xls_sheet)
      xls_sheet[range]
    end

    def range_contains?(range, xls_cell)
      pos                 = xls_cell.pos
      top_left, bot_right = range.split(':').map { |s| Axlsx.name_to_indices(s) }
      pos[0] >= top_left[0] && pos[0] <= bot_right[0] && pos[1] >= top_left[1] && pos[1] <= bot_right[1]
    end
  end
end