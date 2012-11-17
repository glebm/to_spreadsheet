require 'to_spreadsheet/type_from_value'
module ToSpreadsheet
  module Rule
    class DefaultValue < Base
      include ::ToSpreadsheet::TypeFromValue

      def apply(context, sheet)
        default = options
        each_cell context, sheet, selector_type, selector_query do |cell|
          unless cell.value.present? &&
              !([:integer, :float].include?(cell.type) && cell.value.zero?)
            cell.type  = cell_type_from_value(default)
            cell.value = default
          end
        end
      end
    end
  end
end
