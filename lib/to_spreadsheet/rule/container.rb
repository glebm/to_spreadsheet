module ToSpreadsheet
  module Rule
    # Applies children rules to all the matching tables
    class Container < Base
      attr_reader :children
      def initialize(*args)
        super
        @children = []
      end

      def apply(context, package)
        package.workbook.worksheets.each do |sheet|
          table = context.to_xml_node(sheet)
          if applies_to?(context, table)
            children.each { |c| c.apply(context, sheet) }
          end
        end
      end

      def to_s
        "Rules(#{selector_type}, #{selector_query}) [#{children.map(&:to_s)}]"
      end
    end
  end
end
