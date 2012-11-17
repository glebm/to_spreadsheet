module ToSpreadsheet
  module Helpers
    def format_xls(selector = nil, &block)
      context.format_xls selector, &block
    end

    def context
      @context || ToSpreadsheet::Context.global
    end

    def with_context(context, &block)
      context_was = self.context
      @context    = context
      result      = block.call
      @context    = context_was
      result
    end
  end
end