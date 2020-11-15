require_relative './spreadsheet_connection.rb'

class Config
  def initialize(log)
  @spreadsheet = GoogleSpreadsheet.new
 end

 def config
  @results = @spreadsheet.get_spreadsheet_data('12DlPT2Md77ZYRgZFQlPRxa4dPBcNrbsjUiOdwLLeFK4', 'cursos!C2:E')

  return @results
 end

  def read_filters
    @results.map! do |row|
      if row[1].downcase != "todas"
        aula_ids = row[1].split(/ ?, ?/).map{ |x| x.to_i.to_s }
      else
        aula_ids = []
      end
      row[2] ||= ''
      {report_id: row[0], aula_ids: aula_ids, fecha_inicio: row[2].gsub('/', '-')}
    end

    return @results
  end

  def config2
    @results = @spreadsheet.get_spreadsheet_data('12DlPT2Md77ZYRgZFQlPRxa4dPBcNrbsjUiOdwLLeFK4', 'backups!A2:K')
    index = 1
    @results.map! do |r|
      index += 1
      [index] + r
    end

    return @results
 end

  def read_filters2
    @results.map! do |row|
      report_id = row[1]
      path = File.join(row[5..11])

      {row: row[0], report_id: report_id, path: path}
    end
    #@results.delete_if{ |r| r[:path] == '' }

    return @results
  end

  def reportes
    config
    read_filters

    return @results
  end

  def backup
    config2
    read_filters2

    @results
  end
end
