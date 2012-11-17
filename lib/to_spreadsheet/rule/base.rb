require 'to_spreadsheet/selectors'
module ToSpreadsheet
  module Rule
    class Base
      include ::ToSpreadsheet::Selectors
      attr_reader :selector_type, :selector_query, :options

      def initialize(selector_type, selector_query, options)
        @selector_type  = selector_type
        @selector_query = selector_query
        @options        = options
      end

      def applies_to?(context, xml_or_xls_node)
        return true if !selector_type
        node, entity = context.xml_node_and_xls_entity(xml_or_xls_node)
        sheet        = entity.is_a?(::Axlsx::Workbook) ? entity : (entity.respond_to?(:workbook) ? entity.workbook : entity.worksheet.workbook)
        doc          = context.to_xml_node(sheet)
        query_match?(
            selector_type:  selector_type,
            selector_query: selector_query,
            xml_document:   doc,
            xml_node:       node,
            xls_worksheet:  sheet,
            xls_entity:     entity
        )
      end

      def type
        self.class.name.demodulize.underscore.to_sym
      end

      def to_s
        "Rule [#{type}, #{selector_type}, #{selector_query}, #{options}"
      end
    end
  end
end