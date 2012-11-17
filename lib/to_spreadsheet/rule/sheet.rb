module ToSpreadsheet
  module Rule
    class Sheet < Base
      def apply(context, sheet)
        options.each { |k, v|
          if v.is_a?(Hash)
            sub = sheet.send(k)
            v.each do |sub_k, sub_v|
              sub.send :"#{sub_k}=", sub_v
            end
          else
            sheet.send :"#{k}=", v
          end
        }
      end
    end
  end
end
