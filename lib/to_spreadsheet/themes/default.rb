module ToSpreadsheet::Themes
  module Default
    ::ToSpreadsheet.theme :default do
      workbook use_autowidth: true
      sheet page_setup: {
          fit_to_height: 1,
          fit_to_width:  1,
          orientation:   :landscape
      }
      # Set value type based on CSS class
      format 'td,th', lambda { |cell|
        val = cell.value
        case node_from_entity(cell)[:class]
          when /decimal|float/
            cell.type = :float
          when /num|int/
            cell.type = :integer
          when /bool/
            cell.type = :boolean
          # Parse (date)times and dates with Chronic and Date.parse
          when /datetime|time/
            val = Chronic.parse(val)
            if val
              cell.type  = :time
              cell.value = val
            end
          when /date/
            val = (Date.parse(val) rescue val)
            if val.present?
              cell.type  = :date
              cell.value = val
            end
        end
      }
    end
  end
end