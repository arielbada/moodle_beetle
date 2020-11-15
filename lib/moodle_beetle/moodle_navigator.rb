require 'watir'
require 'webdrivers/chromedriver'

class MoodleNavigator

  COORDENADA = 50
  WINDOWS_LIMIT = 3  
  DOWNLOAD_DIR = File.join(File.expand_path(Dir.pwd), 'download')

  def initialize(username, password, log)
    @username = username
    @password = password
    @log = log
    @today = Time.now.strftime('%d/%m/%Y')
    @windows_opened = 0
    FileUtils.rm_rf(DOWNLOAD_DIR)
    load_chromedriver    
  end

  def load_chromedriver
    Webdrivers::Chromedriver.required_version = '86.0'
    chromedriver_path = './ext/webdrivers/chromedriver'
    chromedriver_path += '.exe' unless os_is_linux?
    Selenium::WebDriver::Chrome::Service.driver_path = chromedriver_path
  end

  def download_from_moodle(type, reference_records)  # reference_records = [{:report_id=>"1073", :aula_ids=>[], :fecha_inicio=>"18-02-2018"}, {:report_id=>"1075", :aula_ids=

    @log.info 'navegando la plataforma...'

    if type == 'sited'
      @browser = start_webdriver
      @browser.link(text: "Cuantitativo por Categoría").click!
    end

    reference_records.each do |record|
      case type
      when 'sited'
        generar_sited(record)
      when 'inactividad_docente'
        generar_reporte_docente(record)
      end
      sleep 7
    end

    sleep 1 while (Dir[File.join(DOWNLOAD_DIR,'*','*.xls*'), File.join(DOWNLOAD_DIR, '*.xls*')].count < reference_records.count)# wait to have all reports downloaded

  end

  private

  def generar_sited(record)
    @log.info "bajando reporte #{record[:report_id]}"
    @browser.text_field('name': 'categoria').set(record[:report_id])
    @browser.date_field('name': 'inicio').set(record[:fecha_inicio].gsub('-', '/'))
    @browser.date_field('name': 'fin').set(@today)
    @browser.button(name: 'generar').click!
  end

  def start_webdriver(download_folder = '')
    download_prefs = {
        "directory_upgrade"=> true,
        "prompt_for_download"=> false,
        "default_directory"=> new_download_folder(download_folder)
    }
    options = Selenium::WebDriver::Chrome::Options.new.tap do |o|
      o.add_preference(:download, download_prefs)
      # o.add_preference(:browser, set_download_behavior: { behavior: 'allow' })
      o.add_option(:detach, true)
      o.add_argument('--no-sandbox')
      o.add_argument('--disable-infobars')
      o.add_argument('--disable-browser-side-navigation')
      o.add_argument('--headless') unless DEBUG
      o.add_argument('--disable-gpu')
      o.add_argument('--log-level=3')
    end

    browser = Watir::Browser.new :chrome, options: options
    browser.driver.manage.timeouts.page_load = 90

    browser = login(browser)

    browser
  end
 
  def folder_name_from_id(report_id)
    File.join(DOWNLOAD_DIR, report_id)
  end

  def login(browser)
    unless @login
      @log.info 'logueando usuario...'
      browser.goto('https://plataformaeducativa.santafe.edu.ar/moodle/login/index.php')

      browser.text_field('ng-model': "user.name").set(@username)
      browser.text_field('ng-model': "user.password").set(@password)
      browser.button(class: 'md-btn md-raised green btn-block'.split(' ')).click!

      @login = true
    end

    return browser
  end

  def new_download_folder(folder_name = '')
    folder_name =  File.join(DOWNLOAD_DIR, folder_name)
    FileUtils.mkdir_p(folder_name)

    if os_is_linux?
      folder_name
    else
      folder_name.gsub('/', '\\')
    end
  end

  def os_is_linux?
    RUBY_PLATFORM.match(/linux/i)
  end
end
