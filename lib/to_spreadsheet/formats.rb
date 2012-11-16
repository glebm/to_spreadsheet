module ToSpreadsheet
  # An intrnal container of formats
  class Formats
    attr_accessor :styles_by_type
    attr_writer :sheet_props
    
    # A list of [format_selector, formatting block] by `type`
    # `type` is one of :sheet, :workbook, :range, :column, :row, :css
    def [](type)
      (@styles_by_type ||= {})[type.to_sym] ||= []
    end

    def each
      @styles_by_type.each do |k, v|
        yield(k, v)
      end if @styles_by_type
    end

    # A list of all formatting blocks with type :sheet
    def sheet_props
      self[:sheet].map(&:last).inject({}, &:merge)
    end

    # Workbook props without selectors
    def workbook_props
      self[:workbook].map(&:last).inject({}, &:merge)
    end

    def range_props
      self[:range]
    end

    def column_props
      self[:column]
    end

    def row_props
      self[:row]
    end

    def css_props
      self[:css]
    end

    def derive
      derived = Formats.new
      each { |type, styles| derived[type].concat(styles) }
      derived
    end

    def merge!(other_fmt)
      other_fmt.each { |type, styles| self[type].concat(styles) }
    end

    def inspect
      "Formats(sheet: #@sheet_props, styles: #@styles_by_type)"
    end
  end
end