require 'to_spreadsheet/version'
require 'to_spreadsheet/context'
require 'to_spreadsheet/renderer'

module ToSpreadsheet
  class << self
    def renderer
      @renderer ||= :xlsx
    end

    def theme(name, &formats)
      @themes ||= {}
      if formats
        @themes[name] = formats
      else
        @themes[name]
      end
    end
  end
end

require 'to_spreadsheet/railtie' if defined?(Rails)
require 'to_spreadsheet/themes/default'
ToSpreadsheet::Context.global.format_xls ToSpreadsheet.theme(:default)