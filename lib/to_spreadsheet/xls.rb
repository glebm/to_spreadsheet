module ToSpreadsheet
  require 'spreadsheet'
  module XLS
    extend self

    def to_io(html)
      spreadsheet = Spreadsheet::Workbook.new

      pageheader_fmt = Spreadsheet::Format.new :weight => :bold, :size => 14
      pagesubheader_fmt = Spreadsheet::Format.new :weight => :bold, :size => 13
      tblheader_fmt = Spreadsheet::Format.new :weight => :bold, :size => 12
      th_fmt = Spreadsheet::Format.new :weight => :bold 
      money_fmt = Spreadsheet::Format.new :number_format => '0.00;[red](-0)'
      date_fmt = Spreadsheet::Format.new :number_format => 'MM-DD-YYYY'

      sheet = nil
      row = nil
      colwidths = []
      Nokogiri::HTML::Document.parse(html).css('table').each_with_index do |xml_table, i|
        # Set newtab=1 on a table in the layout where it should be on a new sheet
        if sheet.nil? || xml_table['newtab']
          sheetname = xml_table.css('caption').inner_text.presence || xml_table['name'] || "Sheet #{i + 1}"
          sheet = spreadsheet.create_worksheet(:name => sheetname)
          row = 0
          colwidths = []
        end

        # Either set name="name" or <table><caption>name</caption> to put a header before the table starts
#        tblname = xml_table.css('caption').inner_text.presence || xml_table['name'] || nil
#        if tblname
#          sheet[row, 0] = tblname
#          sheet.row(row).set_format 0, tblheader_fmt
#          sheet.row(row).height = 16
#          row += 1
#        end

        xml_table.css('tr').each do |row_node|
          row_node.css('th,td').each_with_index do |col_node, col|
            sheet[row, col] = typed_node_val(col_node)

            if col_node[:class] =~ /subheader/
              sheet.row(row).height = 17.7
              sheet.row(row).set_format col, pagesubheader_fmt
            elsif col_node[:class] =~ /header/
              sheet.row(row).height = 18.7
              sheet.row(row).set_format col, pageheader_fmt
            end

            if col_node[:class] =~ /date/
              sheet.row(row).set_format col, date_fmt
            elsif col_node[:class] =~ /money/
              sheet.row(row).set_format col, money_fmt
            elsif col_node.name == 'th'
              sheet.row(row).set_format col, th_fmt
            end

            fmt = sheet.row(row).formats[col]
            fmt ||= Spreadsheet::Format.new
            borders = fmt.border

            if col_node[:class] =~ /green/
              fmt.pattern_fg_color = :lime #'#d7e4bc'
              fmt.pattern = 1
            end

            fmt.border = [borders[3], borders[2], true, borders[0]] if col_node[:class] =~ /border_top/
            fmt.border = [borders[3], borders[2], borders[1], true] if col_node[:class] =~ /border_bottom/

            sheet.row(row).set_format col, fmt

            # dynamic column widths
            len = sheet[row, col].to_s.length
            if len > (colwidths[col] || 0)
              colwidths[col] = len
              sheet.column(col).width = len + 2
            end
          end
          row += 1
        end
        row += 1 # extra space between tables on same sheet
      end
      io = StringIO.new
      spreadsheet.write(io)
      io.rewind
      io
    end

    private

    def typed_node_val(node)
      val = node.inner_text
      case node[:class]
        when /decimal|float/
          val.to_f
        when /num|int/
          val.to_i
        when /datetime/
          DateTime.parse(val)
        when /date/
          Date.parse(val)
        when /time/
          Time.parse(val)
        else
          val
      end
    end
  end
end
