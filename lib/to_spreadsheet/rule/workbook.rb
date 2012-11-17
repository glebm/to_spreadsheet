module ToSpreadsheet
  module Rule
    class Workbook < Base
      def apply(context, sheet)
        workbook = sheet.workbook
        options.each { |k, v| workbook.send :"#{k}=", v }
      end
    end
  end
end
