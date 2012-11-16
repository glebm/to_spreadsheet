require 'to_spreadsheet/context/pairing'
require 'to_spreadsheet/formats'

module ToSpreadsheet
  # This is the DSL context for `format_xls`
  # It maintains the current formats set to enable for local and nested `format_xls` blocks
  class Context
    # todo (cleaner code): split features further into modules
    # todo (extensibility): add processing callbacks (internal API)
    include Pairing

    def initialize(wb_options = nil)
      @formats        = []
      @current_format = Formats.new
      workbook wb_options if wb_options
    end

    # Returns a new formats jar for a given sheet
    def formats(sheet)
      format = Formats.new
      @formats.each do |v|
        sel, fmt = v[0], v[1]
        format.merge!(fmt) if selects?(sel, sheet)
      end
      format
    end

    # Check if selector matches a given sheet / cell / row
    def selects?(selector, entity)
      return true if !selector
      type, val = selector[0], selector[1]
      sheet     = entity.is_a?(::Axlsx::Workbook) ? entity : (entity.respond_to?(:workbook) ? entity.workbook : entity.worksheet.workbook)
      doc       = node_from_entity(sheet)
      case type
        when :css
          doc.css(val).include?(node_from_entity(entity))
        when :column
          return false if entity.is_a?(Axlsx::Row)
          entity.index == val if entity.is_a?(Axlsx::Cell)
        when :row
          return entity.index == val if entity.is_a?(Axlsx::Row)
          entity.row.index == val if entity.is_a?(Axlsx::Cell)
        when :range
          if entity.is_a?(Axlsx::Cell)
            pos                 = entity.pos
            top_left, bot_right = val.split(':').map { |s| Axlsx.name_to_indices(s) }
            pos[0] >= top_left[0] && pos[0] <= bot_right[0] && pos[1] >= top_left[1] && pos[1] <= bot_right[1]
          end
      end
    end

    # current format, used internally
    attr_accessor :current_format

    # format_xls 'table.zebra' do
    #   format 'td', lambda { |cell| {b: true} if cell.row.even? }
    # end
    def format_xls(selector = nil, theme = nil, &block)
      selector, theme = nil, selector if selector.is_a?(Proc) && !theme
      add_format(selector, &theme) if theme
      add_format(selector, &block) if block
      self
    end

    # format 'td.b', b: true # bold
    # format column: 0, width: 50
    # format 'A1:C30', b: true
    # Accepted properties: http://rubydoc.info/github/randym/axlsx/Axlsx/Cell
    # column format also accepts Axlsx columnInfo settings
    def format(selector = nil, options)
      options  = options.dup
      selector = extract_selector!(selector, options)
      add selector[0], selector, options
    end

    # sheet 'table.landscape', page_setup: { orientation: landscape }
    def sheet(selector = nil, options)
      options  = options.dup
      selector = extract_selector!(selector, options)
      add :sheet, selector, options
    end

    # default 'td.c', 5
    def default(selector, value)
      options  = {default_value: value}
      selector = extract_selector!(selector, options)
      add selector[0], selector, options
    end

    def add(setting, selector, value)
      @current_format[setting] << [selector.try(:[], 1), value] if selector || value
    end

    def workbook(selector = nil, value)
      add :package, selector, value
    end

    def apply(theme = nil, &block)
      add_format &theme if theme
      add_format &block if block
      self
    end

    def derive
      derived        = dup
      derived.current_format = derived.current_format.derive
      derived
    end

    private

    def add_format(sheet_sel = nil, &block)
      format_was      = @current_format
      @current_format = @current_format.derive
      instance_eval &block
      @formats << [extract_selector!(sheet_sel), @current_format]
      @current_format = format_was
    end

    def extract_selector!(selector, options = {})
      if selector
        if selector =~ /:/ && selector[0].upcase == selector[0]
          return [:range, selector]
        else
          return [:css, selector]
        end
      end
      [:column, :row, :range].each do |key|
        return [key, options.delete(key)] if options.key?(key)
      end
      selector
    end
  end
end