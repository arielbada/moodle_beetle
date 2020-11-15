
require 'roo-xls'
require 'csv'
require 'date'
require 'write_xlsx'
require 'net_http_ssl_fix'
require 'logger'

require_relative './moodle_beetle/moodle_navigator.rb'
require_relative './moodle_beetle/excel_parser.rb'
require_relative './moodle_beetle/multiio.rb'

DEBUG = true

class VVReports

  def initialize
    log_file = File.open("log.txt", "a")
    @log = Logger.new(MultiIO.new(STDOUT, log_file))
    @log.datetime_format = "%Y%m%d %H:%M:%S" 
    @log.info 'iniciando proceso...'    
  end

  def read_reference_records
    puts 'leyendo lista de aulas a descargar...'
    ###
    @reference_records = [{:report_id=>"1073", :aula_ids=>[], :fecha_inicio=>"18-02-2018"}]
    ###
  end

  def generate_output_report
    excel_parser = ExcelParser.new(@log)
    excel_parser.generate!(@process_name, @reference_records)
  end

  def arguments_are_correct?
    if ['inactividad_docente', 'sited'].include?(@process_name)
      @arguments_are_correct = true
    else
      @arguments_are_correct = false
    end
    @arguments_are_correct
  end

  def parse_arguments!
    if ARGV.count != 3
      @log.info 
      @log.info 'Uso: ruby main.rb [NOMBRE_REPORTE] [USUARIO] [PASSWORD]'
      @log.info 'NOMBRE_REPORTE: inactividad_docente | sited'
      @log.info 'USUARIO, PASSWORD: usuario y password para acceder a plataforma'
      return
    end

    @process_name = ARGV[0].downcase
    @user = ARGV[1]
    @password = ARGV[2]
  end

  def run
    parse_arguments!
    return unless arguments_are_correct?

    read_reference_records
    @moodle_navigator = MoodleNavigator.new(@user, @password, @log)
    @moodle_navigator.download_from_moodle(@process_name, @reference_records)
    generate_output_report if ['inactividad_docente', 'sited'].include?(@process_name)
    @log.info 'proceso finalizado'
  end
end

VVReports.new.run
