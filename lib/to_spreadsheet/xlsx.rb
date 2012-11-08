#    Style attributes:
#      bg_color
#      fg_color
#      b = bold
#      i = italic
#      u = underline
#      strike
#      outline
#      shadow
#      charset
#      family
#      font_name
#      hidden
#      locked
#      alignment
#      border = { :style => ?, :color => ?, :edges => [:left, :right, :top, :bottom] }
#
module ToSpreadsheet
  require 'axlsx'
  module XLSX
    extend self

    # Generate an AXLSX Package for the passed HTML table
    def generate(html)
      package = Axlsx::Package.new
      package.use_autowidth = false
      spreadsheet = package.workbook

      sheet = nil
      row = nil
      colwidths = []
      Nokogiri::HTML::Document.parse(html).css('table').each_with_index do |xml_table, i|
        # Set newtab=1 on a table in the layout where it should be on a new sheet
        if sheet.nil? || xml_table['newtab']
          sheetname = xml_table.css('caption').inner_text.presence || xml_table['name'] || "Sheet #{i + 1}"

          page_setup_options = eval ( xml_table['page_setup'] ) rescue nil
          page_setup_options = { fit_to_height: 1, fit_to_width: 1, orientation: :landscape } if page_setup_options.nil?

          sheet = spreadsheet.add_worksheet(:name => sheetname, :page_setup => page_setup_options)
          row = 0
          colwidths = []
        end

        xml_table.css('tr').each do |row_node|
          xlsrow = sheet.add_row

          if ! row_node[:xls_style].nil? && ! row_node[:xls_style].empty?
            row_stylehash = eval( row_node[:xls_style] ) rescue {}
          else
            row_stylehash = {}
          end

          row_node.css('th,td').each_with_index do |col_node, col|

            if ( ! col_node[:xls_style].nil? && ! col_node[:xls_style].empty? ) || ! row_stylehash.empty?
              stylehash = eval( col_node[:xls_style] ) rescue {}
              merged = stylehash.merge(row_stylehash)
              style = spreadsheet.styles.add_style merged
              xlscol = xlsrow.add_cell typed_node_val(col_node), :style => style
            else
              xlscol = xlsrow.add_cell typed_node_val(col_node)
            end

            mywidth = xlscol.value.to_s.length
            mywidth += 3 if spreadsheet.styles.fonts[spreadsheet.styles.cellXfs[xlscol.index].fontId].b rescue false
            if colwidths[xlscol.index].nil? || mywidth> colwidths[xlscol.index]
              colwidths[xlscol.index] = mywidth
            end

          end

          row += 1
        end
        colwidths.each_with_index.map { |len, i| sheet.column_info[i].width = len + 2 }
        row += 1 # extra space between tables on same sheet
      end
      package
    end

    private

    def typed_node_val(node)
      val = val_or_default(node)

      case node[:class]
        when /decimal|float/
          val.to_f
        when /num|int/
          val.to_i
        when /datetime/
          DateTime.parse(val) rescue val
        when /date/
          Date.parse(val) rescue val
        when /time/
          Time.parse(val) rescue val
        else
          val
      end
    end

    def val_or_default(node)
      val = node.inner_text
      if val.blank?
        node['data-default']
      else
        val
      end
    end

  end
end
