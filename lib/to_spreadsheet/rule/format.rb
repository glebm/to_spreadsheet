require 'to_spreadsheet/type_from_value'
require 'set'
module ToSpreadsheet
  module Rule
    class Format < Base
      include ::ToSpreadsheet::TypeFromValue
      def apply(context, sheet)
        wb = sheet.workbook
        case selector_type
          when :css
            css_match selector_query, context.to_xml_node(sheet) do |xml_node| 
              add_and_apply_style wb, context, context.to_xls_entity(xml_node)
            end
          when :row
            sheet.row_style selector_query, options if options.present?
          when :column
            inline_styles = options.except(*COL_INFO_PROPS)
            sheet.col_style selector_query, inline_styles if inline_styles.present?
            apply_col_info sheet.column_info[selector_query]
          when :range
            add_and_apply_style wb, range_match(selector_query, sheet), context
        end
      end

      private
      COL_INFO_PROPS = %w(bestFit collapsed customWidth hidden phonetic width).map(&:to_sym).to_set
      def apply_col_info(col_info)
        return if col_info.nil?
        options.each do |k, v|
          if COL_INFO_PROPS.include?(k)
            col_info.send :"#{k}=", v
          end
        end
      end

      def add_and_apply_style(wb, context, xls_ent)
        # Custom format rule
        # format 'td.sel', lambda { |node| ...}
        if self.options.is_a?(Proc)
          context.instance_exec(xls_ent, &self.options)
          return
        end

        options = self.options.dup
        # Compute Proc rules
        # format 'td.sel', color: lambda {|node| ...}
        options.each do |k, v|
          options[k] = context.instance_exec(xls_ent, &v) if v.is_a?(Proc)
        end

        style = wb.styles.add_style options
        cells = xls_ent.respond_to?(:cells) ? xls_ent.cells : [xls_ent]
        cells.each { |cell| cell.style = style }
      end
    end
  end
end
