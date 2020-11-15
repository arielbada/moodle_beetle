
require 'roo-xls'
require 'csv'
require 'date'
require 'write_xlsx'
require 'net_http_ssl_fix'
require 'logger'

require_relative './moodle_beetle/reportes_config.rb'
require_relative './moodle_beetle/moodle_navigator.rb'
require_relative './moodle_beetle/excel_parser.rb'
require_relative './moodle_beetle/multiio.rb'

DEBUG = false
CONFIG_SHEET = 'https://docs.google.com/spreadsheets/d/12DlPT2Md77ZYRgZFQlPRxa4dPBcNrbsjUiOdwLLeFK4'

class VVReports

  def initialize
    log_file = File.open("log.txt", "a")
    @log = Logger.new(MultiIO.new(STDOUT, log_file))
    @log.datetime_format = "%Y%m%d %H:%M:%S" 
    @log.info 'iniciando proceso...'
  end

  def read_config
    @log.info "leyendo configuracion desde #{CONFIG_SHEET}"    
    @config = Config.new(@log).reportes
    
    @log.info "#{@config.count} filtros cargados"
  end

  def generate_output_report
    excel_parser = ExcelParser.new(@log)
    excel_parser.generate!(@process_name, @config)
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

    read_config
    @moodle_navigator = MoodleNavigator.new(@user, @password, @log)
    @moodle_navigator.download_from_moodle(@process_name, @config)
    generate_output_report if ['inactividad_docente', 'sited'].include?(@process_name)
    @log.info 'proceso finalizado'
  end
end

VVReports.new.run
