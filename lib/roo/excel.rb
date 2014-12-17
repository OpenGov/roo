require 'spreadsheet'

class Worksheet < ::Spreadsheet::Excel::Worksheet
  def each_with_index(*args)
    return enum_for(:each_with_index, *args) unless block_given?

    index = 0
    each(*args) do |v|
      yield v, index
      index += 1
    end
  end
end

# ruby-spreadsheet has a font object so we're extending it
# with our own functionality but still providing full access
# to the user for other font information
class ::Spreadsheet::Font
  def bold?(*args)
    #From ruby-spreadsheet doc: 100 <= weight <= 1000, bold => 700, normal => 400
    weight == 700
  end

  def italic?
    italic
  end

  def underline?
    underline != :none
  end
end

# Class for handling Excel-Spreadsheets
class Roo::Excel < Roo::Base
  FORMULAS_MESSAGE = 'the spreadsheet gem does not support forumulas, so roo can not.'

  attr_reader :workbook

  # Creates a new Excel spreadsheet object.
  # Parameter packed: :zip - File is a zip-file
  # Parameter mode: 'rb+' - Mode of the xls file reader
  # Parameter input_encoding: Encoding:: - Encoding of the embeded binary data (default nil)
  # Parameter output_encoding: Encoding:: - Encoding of the internal string structures
  def initialize(filename, options = {}, deprecated_file_warning = nil)
    unless Hash === options
      warn 'Supplying `packed` or `file_warning` as separate arguments to `Roo::Excel.new` is deprecated. Use an options hash instead.'
      options = { packed: options }
    end

    packed = options[:packed]
    file_warning = options[:file_warning] || deprecated_file_warning || :error
    mode = options[:mode] || 'rb+'
    @input_encoding = options[:input_encoding]
    @output_encoding = options[:output_encoding] || Encoding::UTF_8

    file_type_check(filename, '.xls', 'an Excel', file_warning, packed)
    make_tmpdir do |tmpdir|
      filename = download_uri(filename, tmpdir) if uri?(filename)
      filename = open_from_stream(filename[7..-1], tmpdir) if filename.is_a?(::String) && filename.start_with?('stream:')
      filename = unzip(filename, tmpdir) if packed == :zip

      @filename = filename
      unless File.file?(@filename)
        fail IOError, "file #{@filename} does not exist"
      end
      @workbook = ::Spreadsheet.open(filename, mode)
    end
    super(filename, options)
    @formula = {}
    @fonts = {}
  end

  def worksheets
    @worksheets ||= @workbook.worksheets
  end

  def encoding=(codepage)
    @workbook.encoding = codepage
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    worksheets.collect {|worksheet| normalize_string(worksheet.name)}
  end

  # this method lets you find the worksheet with the most data
  def longest_sheet
    sheet(worksheets.inject {|m, o|
      o.row_count > m.row_count ? o : m
    }.name)
  end

  # Iterates through each row for of a particular sheet
  # Does not pad past the last present cell
  #
  # options[:sheet] can be used to specify a sheet other
  # than the default
  # options[:max_rows] break when parsed max rows
  # options[:strip_nils] removes trailing nils from rows
  def each_row(options = {})
    return enum_for(:each_row, options) unless block_given?

    sheet = options[:sheet] || @default_sheet
    validate_sheet!(sheet)

    cell_sheet = @cell[sheet]
    build_rows = cell_sheet.empty?
    worksheet = @workbook.worksheet(sheet_no(sheet))
    worksheet.each_with_index(0) do |xml_row, row_index|
      row = Array.new(xml_row.size)
      (0...xml_row.size).each do |cell_index|
        key = [row_index + 1, cell_index + 1]
        # Grab our saved values if the cell has already been parsed
        if build_rows
          value_type, v = read_cell(xml_row, cell_index)
          font = xml_row.format(cell_index).font
          value = set_cell_values(sheet, key[0], key[1], 0, v, value_type, nil, nil, font)
        else
          value = row[cell_index] = cell_sheet.fetch(key, nil)
        end
        row[cell_index] = value
      end

      row.pop while options[:strip_nils] && row.last.nil? && !row.empty?
      yield row
      break if options[:max_rows] && row_index + 1 == options[:max_rows]
    end
  end

  # returns the content of a cell. The upper left corner is (1,1) or ('A',1)
  def cell(row, col, sheet = nil)
    sheet ||= @default_sheet
    read_cells(sheet)

    @cell[sheet].fetch(normalize(row, col).to_a, nil)
  end

  # returns the type of a cell:
  # * :float
  # * :string,
  # * :date
  # * :percentage
  # * :formula
  # * :time
  # * :datetime
  def celltype(row, col, sheet = nil)
    sheet ||= @default_sheet
    read_cells(sheet)

    @cell_type[sheet].fetch(normalize(row, col).to_a, nil)
  end

  # returns NO formula in excel spreadsheets
  def formula(_row, _col, _sheet = nil)
    raise NotImplementedError, FORMULAS_MESSAGE
  end
  alias_method :formula?, :formula

  # returns NO formulas in excel spreadsheets
  def formulas(_sheet=nil)
    raise NotImplementedError, FORMULAS_MESSAGE
  end

  # Given a cell, return the cell's font
  def font(row, col, sheet = nil)
    sheet ||= @default_sheet
    read_cells(sheet)

    @fonts[sheet][normalize(row,col).to_a]
  end

  # shows the internal representation of all cells
  # for debugging purposes if debug is truthy
  def to_s(sheet = nil, debug = nil)
    sheet ||= @default_sheet
    # only print the whole thing on debug
    if debug
      read_cells(sheet)
      @cell[sheet].inspect
    else
      super()
    end
  end

  def inspect(sheet = nil, debug = nil)
    to_s(sheet, debug)
  end

  # check if default_sheet was set and exists in sheets-array
  def validate_sheet!(sheet)
    super
    # establish our sheet lookups so we don't do this each cell assignment
    @cell_type[sheet] ||= {}
    @formula[sheet] ||= {}
    @cell[sheet] ||= {}
    @fonts[sheet] ||= {}
  end

  private

  # converts name of a sheet to index (0,1,2,..)
  def sheet_no(name)
    return name - 1 if name.is_a?(Fixnum)
    worksheets.each_with_index do |worksheet, index|
      return index if name == normalize_string(worksheet.name)
    end
    raise StandardError, "sheet '#{name}' not found"
  end

  # copies the input if any encoding changes occur
  def normalize_string(str)
    if str.respond_to?(:encoding)
      if @input_encoding
        force_value = String.new(str).force_encoding(@input_encoding)
        str = force_value if force_value.valid_encoding?
      end
      str = str.encode(@output_encoding, 'binary', undef: :replace) if @output_encoding && str.encoding != @output_encoding
    end

    str
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet, row, col, i, v, value_type, formula, _tr, font)
    # key = "#{y}, #{x+i}"
    key = [row, col + i]
    if formula
      @formula[sheet][key] = formula
      value_type = :formula
    end
    @cell_type[sheet][key] = value_type
    @fonts[sheet][key] = font
    @cell[sheet][key] = v
  end

  # read all cells in the selected sheet
  def read_cells(sheet = nil)
    sheet ||= @default_sheet
    return if @cells_read[sheet] && !@cells_read[sheet].empty?
    each_row(sheet: sheet).each { |_row| }

    @cells_read[sheet]
  end

  # Get the contents of a cell, accounting for the
  # way formula stores the value
  def read_cell_content(row, idx)
    cell = row.at(idx)
    if cell.class == ::Spreadsheet::Formula
      # Formulas oftentimes lose type information
      cell = cell.value.to_f rescue cell.value
      cell = cell.to_i if (cell.to_f % 1).zero? rescue cell
    end

    cell
  end

  # Test the cell to see if it's a valid date/time.
  def date_or_time?(row, idx)
    row.format(idx).date_or_time? && Float(read_cell_content(row, idx)) > 0 rescue false
  end

  # Read the date-time cell and convert to,
  # the date-time values for Roo
  def read_cell_date_or_time(row, idx)
    cell = read_cell_content(row, idx)
    cell = cell.to_s.to_f
    if cell < 1.0
      value_type = :time
      f = cell * 24.0 * 60.0 * 60.0
      secs = f.round
      h = (secs / 3600.0).floor
      secs = secs - 3600 * h
      m = (secs / 60.0).floor
      secs = secs - 60 * m
      s = secs
      value = h * 3600 + m * 60 + s
    else
      datetime = row.send(:_datetime, cell)

      if datetime.hour != 0 ||
          datetime.min != 0 ||
          datetime.sec != 0
        value_type = :datetime
        value = datetime
      else
        value_type = :date
        value = row.send(:_date, cell)
      end
    end

    [value_type, value]
  end

  # Read the cell and based on the class,
  # return the values for Roo
  def read_cell(row, idx)
    return read_cell_date_or_time(row, idx) if date_or_time?(row, idx)

    cell = read_cell_content(row, idx)
    case cell
    when Integer, Fixnum, Bignum
      value_type = :float
      value = cell.to_i
    when Float
      value_type = :float
      value = cell.to_f
    when String, TrueClass, FalseClass
      value_type = :string
      value = normalize_string(cell.to_s)
    when ::Spreadsheet::Link
      value_type = :link
      value = cell
    else
      value_type = cell.class.to_s.downcase.to_sym
      value = nil
    end # case

    [value_type, value]
  end

end
