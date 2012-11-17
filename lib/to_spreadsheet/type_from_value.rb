module ToSpreadsheet
  module TypeFromValue
    def cell_type_from_value(v)
      if v.is_a?(Date)
        :date
      elsif v.is_a?(Time)
        :time
      elsif v.is_a?(TrueClass) || v.is_a?(FalseClass)
        :boolean
      elsif v.to_s.match(/\A[+-]?\d+?\Z/) #numeric
        :integer
      elsif v.to_s.match(/\A[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?\Z/) #float
        :float
      else
        :string
      end
    end
  end
end