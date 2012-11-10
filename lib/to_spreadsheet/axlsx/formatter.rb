module ToSpreadsheet
  module Axlsx
    module Formatter
      COL_INFO_PROPS = %w(bestFit collapsed customWidth hidden phonetic width).map(&:to_sym)

      def apply_formats(package, context)
        package.workbook.worksheets.each do |sheet|
          fmt = context.formats(sheet)
          fmt.workbook_props.each do |prop, value|
            package.send(:"#{prop}=", value)
          end
          fmt.sheet_props.each do |set, props|
            apply_props sheet.send(set), props, context
          end
          fmt.column_props.each do |v|
            idx, props = v[0], v[1]
            apply_props sheet.column_info[0], props.slice(*COL_INFO_PROPS), context
            props = props.except(*COL_INFO_PROPS)
            sheet.col_style idx, props if props.present?
          end
          fmt.row_props.each do |v|
            idx, props = v[0], v[1]
            sheet.row_style idx, props
          end
          fmt.range_props.each do |v|
            range, props = v[0], v[1]
            apply_props sheet[range], props, context
          end
          fmt.css_props.each do |v|
            css_sel, props = v[0], v[1]
            context.node_from_entity(sheet).css(css_sel).each do |node|
              apply_props context.entity_from_node(node), props, context
            end
          end
        end
      end

      private
      def apply_props(obj, props, context)
        if props.is_a?(Proc)
          context.instance_exec(obj, &props)
          return
        end

        props = props.dup
        props.each do |k, v|
          props[k] = context.instance_exec(obj, &v) if v.is_a?(Proc)
        end

        props.each do |k, v|
          next if v.nil?
          if k == :default_value
            unless obj.value.present? && !([:integer, :float].include?(obj.type) && obj.value.zero?)
              obj.type  = cell_type_from_value(v)
              obj.value = v
            end
          else
            setter = :"#{k}="
            obj.send setter, v if obj.respond_to?(setter)
            if obj.respond_to?(:cells)
              obj.cells.each do |cell|
                cell.send setter, v if cell.respond_to?(setter)
              end
            end
          end
        end
      end

      def cell_type_from_value(v)
        if v.is_a?(Date)
          :date
        elsif v.is_a?(Time)
          :time
        elsif v.is_a?(TrueClass) || v.is_a?(FalseClass)
          :boolean
        elsif v.to_s.match(/\A[+-]?\d+?\Z/) #numeric
          :integer
        elsif v.to_s.match(/\A[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?\Z/) #float
          :float
        else
          :string
        end
      end
    end
  end
end