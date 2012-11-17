require 'to_spreadsheet/type_from_value'
require 'set'
module ToSpreadsheet
  module Rule
    class Format < Base
      include ::ToSpreadsheet::TypeFromValue
      def apply(context, sheet)
        case selector_type
          when :css
            css_match selector_query, context.to_xml_node(sheet) do |xml_node|
              apply_inline_styles context, context.to_xls_entity(xml_node)
            end
          when :row
            sheet.row_style selector_query, options if options.present?
          when :column
            inline_styles = options.except(*COL_INFO_PROPS)
            sheet.col_style selector_query, inline_styles if inline_styles.present?
            apply_col_info sheet.column_info[selector_query]
          when :range
            apply_inline_styles range_match(selector_type, sheet), context
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

      def apply_inline_styles(context, xls_ent)
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

        # Apply inline styles
        options.each do |k, v|
          next if v.nil?
          setter = :"#{k}="
          xls_ent.send setter, v if xls_ent.respond_to?(setter)
          if xls_ent.respond_to?(:cells)
            xls_ent.cells.each do |cell|
              cell.send setter, v if cell.respond_to?(setter)
            end
          end
        end
      end
    end
  end
end