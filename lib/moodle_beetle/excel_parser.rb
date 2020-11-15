
class ExcelParser
  CSV_OUTPUT = "total.csv"
  BOM = "\xEF\xBB\xBF" #Byte Order Mark
  AULA_EXCEPT = Regexp.union(/00$/)
  DNI_EXCEPT = '01330975'
  CSV_DELIMITER = ';'
  DOWNLOAD_DIR = File.join(Dir.pwd, 'descargas')

  def initialize(log)
    @log = log
  end

  def generate!(process_name, reference_records)
    @aula_filters = reference_records  # @aula_filters = [{:report_id=>"934", :aula_ids=>["4926"]}, ...
    case process_name
    when 'sited'
      run_sited_output
    when 'inactividad_docente'
      run_inactividad_docente_output
    end
  end

  def parse_sited_excel(excel_path)
    initialize_excel(excel_path)

    sheet_parsed = @sheet.parse(aula_id: /ID Aula/, aula:/^Aula/, dni: /Documento/, apellido: /Apellido/, nombre: /Nombre/, email: /Email/, localidad: /Localidad/, ultimo_acceso: /Último acceso/)    
    sheet_parsed.each{ |record| record[:report_id] = excel_path.match(/(\d+) ?\.xls/)[1] }

    sheet_parsed
  end
  
  def initialize_excel(excel_path)
    begin
      if File.extname(excel_path) == '.xls'
        @excel = Roo::Spreadsheet.open(excel_path, extension: :xls)
      else
        @excel = Roo::Spreadsheet.open(excel_path, extension: :xlsx)
      end
      @sheet = @excel.sheet(0)
    rescue StandardError => e
      @log.info "error abriendo el archivo excel #{excel_path}" + "\n" + e.to_s
      exit
    end
  end

  def sanitize_data
    r = []
    @register_colector.select!{ |reg| !reg[:aula].match(AULA_EXCEPT) }  # deletes aulas 00
    @register_colector.select!{ |reg| !reg[:dni].to_s.match(DNI_EXCEPT) }  # deletes DNI de tutor virtual 01330975

    begin
      @register_colector.each do |row|
        r = row
        row[:dni] = row[:dni].to_i
        row[:modulo] = row[:aula].match(/([A-Z][A-Z]M\d+)/)[1]
        row[:ultimo_acceso] = row[:ultimo_acceso].sub(/(\d+\/\d+\/\d+) [\d\:]+$/, '\1')
        if @type == 'sited'
          row[:aula_id] = row[:aula_id].to_s.sub(/\.0$/, '')
          row[:aula_n] = row[:aula].match(/Aula ?(\d+)/)[1].to_i.to_s
          row.tap { |r| r.delete(:aula) }
        end
      end
    rescue StandardError => e
      p 'error sanitizing data'
      p r
      p e
      exit
    end
  end

  def write_csv
    File.open(CSV_OUTPUT, "w:UTF-8") do |f|
      csv_content = CSV.generate({:col_sep => CSV_DELIMITER}) do |csv|
        @register_colector.each do |row|  # {:aula_id=>"5321", :dni=>27359501, :apellido=>"Moraz Sengel", :nombre=>"Mariana Paola", :email=>"marianamorazsengel@gmail.com", :localidad=>"Tostado", :ultimo_acceso=>"11/03/2018", :report_id=>"1073", :aula_n=>"1", :modulo=>"COM03"}
          #csv << row.values
          next if row[:aula_n] == "0"
          if @type == 'sited'
            csv << [row[:dni], row[:apellido], row[:nombre], row[:email], row[:localidad], row[:ultimo_acceso], row[:aula_n], row[:modulo], row[:report_id]]
          end
        end
      end
      f.puts BOM + csv_content
    end
  end

  def write_xlsx
    workbook = WriteXLSX.new('total.xlsx')    
    workbook.close
  end

  def filter_records
    @new_colection = []

    @aula_filters.each do |filter|
      if filter[:aula_ids].count > 0
        @log.info "filtrando reporte: #{filter[:report_id].to_s}  -  aulas: #{filter[:aula_ids].join(", ")}"
        @register_colector.delete_if { |x| x[:report_id] == filter[:report_id] and !filter[:aula_ids].include?(x[:aula_n]) }
      end
    end

    @register_colector
  end

  def run_sited_output
    @excel_file_path = Dir[File.join('descargas', '*.xls')]
    @register_colector = []
    @excel_file_path.each do |excel_path|
      @register_colector += parse_sited_excel(excel_path)
    end
    @log.info "#{@register_colector.count} registros totales"
    sanitize_data
    @log.info "#{@register_colector.count} registros después de limpieza"
    filter_records
    @log.info "#{@register_colector.count} registros después de filtrados"

    write_csv
  end 
end
