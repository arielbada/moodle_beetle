require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'
EXECUTION_PATH = File.expand_path('..', File.dirname(__FILE__))

class GoogleSpreadsheet
  OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'.freeze
  APPLICATION_NAME = 'Google Sheets API Ruby'.freeze
  CLIENT_SECRETS_PATH = File.join(EXECUTION_PATH,'client_secret.json').freeze
  CREDENTIALS_PATH = File.join(EXECUTION_PATH,'token.yaml').freeze
  SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS

  def initialize
    # Initialize the API
    @service = Google::Apis::SheetsV4::SheetsService.new
    @service.client_options.application_name = APPLICATION_NAME
    @service.authorization = authorize
  end

  ##
  # Ensure valid credentials, either by restoring from the saved credentials
  # files or intitiating an OAuth2 authorization. If authorization is required,
  # the user's default browser will be launched to approve the request.
  #
  # @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials  
  def authorize
    client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
    token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
    authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
    user_id = 'default'
    credentials = authorizer.get_credentials(user_id)
    if credentials.nil?
    url = authorizer.get_authorization_url(base_url: OOB_URI)
    puts 'Abre la siguiente URL en un navegador e ingresa el siguiente ' \
       'código resultante luego de la authorization:\n' + url
    code = STDIN.gets.chomp
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
    end
    credentials
  end

  def get_spreadsheet_data(spreadsheet_id, range) # '12DlPT2Md77ZYRgZFQlPRxa4dPBcNrbsjUiOdwLLeFK4', 'finance!C2:E'
    response = @service.get_spreadsheet_values(spreadsheet_id, range)
    if response.values.nil? || response.values.empty?
      return false
    else
      return response.values
    end
  end

  def write_cell(spreadsheet_id, range, value) # '12DlPT2Md77ZYRgZFQlPRxa4dPBcNrbsjUiOdwLLeFK4', 'finance!C2:E', "lala"

    value_range_object = Google::Apis::SheetsV4::ValueRange.new(values: value)
    response = @service.update_spreadsheet_value(spreadsheet_id, range, value_range_object, value_input_option: 'RAW')
    if response.updated_range.empty?
      puts 'No data was updated.'
      return false
    else
      return response.updated_range
    end
  end

end # Class
