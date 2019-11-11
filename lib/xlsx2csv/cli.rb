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
        sheetdelimiter: '--------',
      }
      parser.parse!(argv)
      @xlsxfile = argv.shift
      @output = argv.shift
    end

    def parser
      opt = OptionParser.new
      opt.on('-a', '--all', TrueClass, 'export all sheets') do |v|
        self.options[:all] = v
      end
      opt.on('-s[=SHEETID]', '--sheet[=SHEETID]', Integer,
             'sheet number to convert') do |v|
        self.options[:sheet] = v
      end
      opt.on('-n[=SHEETNAME]', '--sheetname[=SHEETNAME]', String,
             'sheet name to convert') do |v|
        self.options[:sheetname] = v
      end
      delimiter_desc = <<~EOS
        delimiter - columns delimiter in csv, 'tab' or 'x09'
                                             for a tab (default: comma ',')
      EOS
      opt.on('-d[=DELIMITER]', '--delimiter[=DELIMITER]', String,
             delimiter_desc) do |v|
        self.options[:delimiter] = v
      end
      lineterminator_desc = <<~EOS
        line terminator - lines terminator in csv, '\\n' '\\r\\n'
                                             or '\\r' (default: \\r\\n)
      EOS
      opt.on('-l[=LINETERMINATOR]', '--lineterminator[=LINETERMINATOR]', String,
             lineterminator_desc) do |v|
        self.options[:lineterminator] = v
      end
      opt.on('-f[=DATEFORMAT]', '--dateformat[=DATEFORMAT]', String,
             'override date/time format (ex. %Y/%m/%d)') do |v|
        self.options[:dateformat] = v
      end
      opt.on('--floatformat[=FLOATFORMAT]', String,
             'override float format (ex. %.15f)') do |v|
        self.options[:floatformat] = v
      end
      opt.on('-i', '--ignoreempty', TrueClass, 'skip empty lines') do |v|
        self.options[:ignoreempty] = v
      end
      sheetdelimiter_desc = <<~EOS
        sheet delimiter used to separate sheets, pass '' if
                                             you do not need delimiter, or 'x07' or '\\f' for form
                                             feed (default: '#{options[:sheetdelimiter]}')
      EOS
      opt.on('-p[=SHEETDELIMITER]', '--sheetdelimiter[=SHEETDELIMITER]', String,
             sheetdelimiter_desc) do |v|
        self.options[:sheetdelimiter] = v
      end
      opt
    end

    def execute
      if @xlsxfile.nil?
        puts parser.help
        exit 1
      end
      file = Roo::Spreadsheet.open(@xlsxfile)
      sheets = []
      if options[:all]
        sheets = file.sheets.map { |sheet| file.sheet(sheet) }
      elsif options[:sheetname]
        sheets << file.sheet(options[:sheetname])
      elsif options[:sheet]
        sheets << file.sheet(options[:sheet])
      else
        sheets << file.sheet(0)
      end
      cont = CSV.generate(col_sep: options[:delimiter],
                          row_sep: options[:lineterminator]) do |csv|
        sheets.each_with_index do |sheet, index|
          sheet.each_row_streaming do |row|
            values = []
            row.each do |cell|
              values << cell.value
              # convert float, date etc...
            end
            csv << values
          end
          csv << [options[:sheetdelimiter]] unless sheets[index + 1].nil?
        end
      end
      if @output
        File.write(@output, cont)
      else
        puts cont
      end
    end
  end
end
