module ToSpreadsheet
  module Rails
    module ViewHelpers
      def format_xls(selector = nil, &block)
        ctx = ToSpreadsheet::Context.current
        return unless ctx
        ctx.format_xls selector, &block
      end
    end
  end
end