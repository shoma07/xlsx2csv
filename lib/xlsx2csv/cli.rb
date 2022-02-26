# frozen_string_literal: true

module Xlsx2csv
  class CLI
    attr_accessor :options

    class << self
      def execute(argv)
        new(argv).execute
      end
    end

    def initialize(argv)
      self.options = {
        all: false,
        sheet: 0,
        sheetname: nil,
        delimiter: ',',
        lineterminator: "\r\n",
        dateformat: nil,
        floatformat: nil,
        ignoreempty: false,
        sheetdelimiter: '--------'
      }
      parser.parse!(argv)
      @xlsxfile = argv.shift
      @output = argv.shift
    end

    def parser
      opt = OptionParser.new
      opt.on('-a', '--all', TrueClass, 'export all sheets') do |v|
        options[:all] = v
      end
      opt.on('-s[=SHEETID]', '--sheet[=SHEETID]', Integer, 'sheet number to convert') do |v|
        options[:sheet] = v
      end
      opt.on('-n[=SHEETNAME]', '--sheetname[=SHEETNAME]', String,
             'sheet name to convert') do |v|
        options[:sheetname] = v
      end
      delimiter_description = <<~DESCRIPTION
        delimiter - columns delimiter in csv, 'tab' or 'x09'
                                             for a tab (default: comma ',')
      DESCRIPTION
      opt.on('-d[=DELIMITER]', '--delimiter[=DELIMITER]', String, delimiter_description) do |v|
        options[:delimiter] = v
      end
      lineterminator_desc = <<~DESCRIPTION
        line terminator - lines terminator in csv, '\\n' '\\r\\n'
                                             or '\\r' (default: \\r\\n)
      DESCRIPTION
      opt.on('-l[=LINETERMINATOR]', '--lineterminator[=LINETERMINATOR]', String,
             lineterminator_desc) do |v|
        options[:lineterminator] = v
      end
      opt.on('-f[=DATEFORMAT]', '--dateformat[=DATEFORMAT]', String,
             'override date/time format (ex. %Y/%m/%d)') do |v|
        options[:dateformat] = v
      end
      opt.on('--floatformat[=FLOATFORMAT]', String,
             'override float format (ex. %.15f)') do |v|
        options[:floatformat] = v
      end
      opt.on('-i', '--ignoreempty', TrueClass, 'skip empty lines') do |v|
        options[:ignoreempty] = v
      end
      sheetdelimiter_desc = <<~DESCRIPTION
        sheet delimiter used to separate sheets, pass '' if
                                             you do not need delimiter, or 'x07' or '\\f' for form
                                             feed (default: '#{options[:sheetdelimiter]}')
      DESCRIPTION
      opt.on('-p[=SHEETDELIMITER]', '--sheetdelimiter[=SHEETDELIMITER]', String, sheetdelimiter_desc) do |v|
        options[:sheetdelimiter] = v
      end
      opt
    end

    def execute
      puts(parser.help) && exit(1) if @xlsxfile.nil?

      sheets = get_sheets(@xlsxfile)
      cont = CSV.generate(col_sep: options[:delimiter], row_sep: options[:lineterminator]) do |csv|
        sheets.each_with_index do |sheet, index|
          sheet.each_row_streaming { |row| csv << row.map(&:value) }
          csv << [options[:sheetdelimiter]] unless sheets[index + 1].nil?
        end
      end
      @output ? File.write(@output, cont) : puts(cont)
    end

    # @param xlsxfile [String]
    # @return [Array]
    def get_sheets(xlsxfile)
      file = Roo::Spreadsheet.open(xlsxfile).then do |file|
      return file.sheets.map { file.sheet(_1) } if options[:all]
      return [file.sheet(options[:sheetname])] if options[:sheetname]
      return [file.sheet(options[:sheet])] if options[:sheet]

      [file.sheet(0)]
    end
  end
end
