require 'to_spreadsheet/context/pairing'
require 'to_spreadsheet/rule'
require 'to_spreadsheet/rule/base'
require 'to_spreadsheet/rule/container'
require 'to_spreadsheet/rule/format'
require 'to_spreadsheet/rule/default_value'
require 'to_spreadsheet/rule/sheet'
require 'to_spreadsheet/rule/workbook'

module ToSpreadsheet
  # This is the DSL context
  class Context
    include Pairing
    attr_accessor :rules

    class << self
      def global
        @global ||= new
      end

      def current
        Thread.current[:_to_spreadsheet_ctx]
      end

      def current=(ctx)
        Thread.current[:_to_spreadsheet_ctx] = ctx
      end

      def with_context(ctx, &block)
        old = current
        self.current = ctx
        r = block.call(ctx)
        self.current = old
        r
      end
    end

    def initialize(wb_options = nil)
      @rules = []
      workbook wb_options if wb_options
    end

    # Examples:
    #   format_xls 'table.zebra' do
    #     format 'td', lambda { |cell| {b: true} if cell.row.even? }
    #   end
    #   format_xls ToSpreadsheet.theme(:a_theme)
    #   format_xls 'table.zebra', ToSpreadsheet.theme(:zebra)
    def format_xls(selector = nil, theme = nil, &block)
      selector, theme = nil, selector if selector.is_a?(Proc) && !theme
      process_dsl(selector, &theme) if theme
      process_dsl(selector, &block) if block
      self
    end

    def process_dsl(selector, &block)
      @rule_container = add_rule :container, *selector_query(selector)
      instance_eval(&block)
      @rule_container = nil
    end

    def workbook(selector = nil, value)
      add_rule :workbook, *selector_query(selector), value
    end

    # format 'td.b', b: true # bold
    # format column: 0, width: 50
    # format 'A1:C30', b: true
    # Accepted properties: http://rubydoc.info/github/randym/axlsx/Axlsx/Cell
    # column format also accepts Axlsx columnInfo settings
    def format(selector = nil, options)
      options  = options.dup
      selector = selector_query(selector, options)
      add_rule :format, *selector, options
    end

    # sheet 'table.landscape', page_setup: { orientation: landscape }
    def sheet(selector = nil, options)
      options  = options.dup
      selector = selector_query(selector, options)
      add_rule :sheet, *selector, options
    end

    # default 'td.c', 5
    def default(selector, value)
      selector = selector_query(selector)
      add_rule :default_value, *selector, value
    end

    def add_rule(rule_type, selector_type, selector_value, options = {})
      rule = ToSpreadsheet::Rule.make(rule_type, selector_type, selector_value, options)
      if @rule_container
        @rule_container.children << rule
      else
        @rules << rule
      end
      rule
    end

    # A new context
    def merge(other_context)
      ctx       = Context.new()
      ctx.rules = rules + other_context.rules
      ctx
    end

    private

    # Extract selector query from DSL arguments
    #
    # Figures out text type:
    #  selector_query('td.num') # [:css, "td.num"]
    #  selector_query('A0:B5') # [:range, "A0:B5"]
    #
    # If text is nil, extracts first of row, range, and css keys
    #  selector_query(nil, {column: 0}] # [:column, 0]
    def selector_query(text, opts = {})
      if text
        if text =~ /:/ && text[0].upcase == text[0]
          return [:range, text]
        else
          return [:css, text]
        end
      end
      key = [:column, :row, :range].detect { |key| opts.key?(key) }
      return [key, opts.delete(key)] if key
      [nil, nil]
    end
  end
end