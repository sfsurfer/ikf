require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/installed_app'
require 'google/api_client/auth/storage'
require 'google/api_client/auth/storages/file_store'
require 'fileutils'
require 'spreadsheet'
require 'nokogiri'
require 'open-uri'

Spreadsheet.client_encoding = 'UTF-8'

APPLICATION_NAME = 'Gmail API Ruby Quickstart'
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials',
                             "gmail-ruby-quickstart.json")
SCOPE = 'https://www.googleapis.com/auth/gmail.readonly'

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization request via InstalledAppFlow.
# If authorization is required, the user's default browser will be launched
# to approve the request.
#
# @return [Signet::OAuth2::Client] OAuth2 credentials
def authorize
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))

  file_store = Google::APIClient::FileStore.new(CREDENTIALS_PATH)
  storage = Google::APIClient::Storage.new(file_store)
  auth = storage.authorize

  if auth.nil? || (auth.expired? && auth.refresh_token.nil?)
    app_info = Google::APIClient::ClientSecrets.load(CLIENT_SECRETS_PATH)
    flow = Google::APIClient::InstalledAppFlow.new({
      :client_id => app_info.client_id,
      :client_secret => app_info.client_secret,
      :scope => SCOPE})
    auth = flow.authorize(storage)
    puts "Credentials saved to #{CREDENTIALS_PATH}" unless auth.nil?
  end
  auth
end

# Initialize the API
client = Google::APIClient.new(:application_name => APPLICATION_NAME)
client.authorization = authorize
gmail_api = client.discovered_api('gmail', 'v1')

# Show the user's labels
results = client.execute!(
  :api_method => gmail_api.users.messages.list,
  :parameters => { 
      :userId => 'me',
      :includeSpamTrash => false,
      :q => 'newer_than:2d subject:"I Know First -Daily Forecast" from:iknowfirst@iknowfirst.com'
 })
#p results.data
puts "Messages:"
puts "No messages found" if results.data.messages.empty?

id = results.data.messages[0].id
#p "message id = #{id}"

r2 = client.execute!(
  :api_method => gmail_api.users.messages.get,
  :parameters => {
    :userId => 'me',
    :id => id
  })


attachment_id = r2.data.payload.parts[1].body.attachmentId
#p "attachment_id = #{attachment_id}"


r3 =  client.execute!(
  :api_method => gmail_api.users.messages.attachments.get,
  :parameters => {
    :userId => 'me',
    :messageId => id,
    :id => attachment_id
  })

attachment_data = JSON.parse(r3.response.body)
#attachment_data = r3.response.body
#attachment_json = JSON.parse(attachment_data.to_json())
#mime_data = Base64.decode64(attachment_json.data['data'])

dest = open('./test.xls','w')
#dest.write(attachment_data['data'])
#dest.write(mime_data)
dest.write(Base64.urlsafe_decode64(attachment_data['data']))
dest.close

longs={}
shorts={}

Spreadsheet.open('./test.xls') do |ikf|
  sheet = ikf.worksheet('1-3-12months')
  sheet.each_with_index do |row,i|
    row.each_with_index do |cell,j|
      break if j > 5
      next if cell.nil?
      next unless cell.is_a? Numeric
      next unless cell > 60 || cell < -60
      tkr = sheet.row(i-1)[j]
      longs[tkr]  = cell if cell > 60
      shorts[tkr] = cell if cell < 60
#      p "Ticker => #{tkr} -> #{cell}"
    end
  end
end
#p "Longs = #{longs}"
#p "Shorts = #{shorts}"
@long_details={}
longs.keys.each do |sym|
  hp_url="http://finance.yahoo.com/q/hp?s=#{sym}+Historical+Prices"
  page = Nokogiri::HTML(open(hp_url))
  rows = page.xpath("//table[contains(@class,'yfnc_datamodoutline1')]/tr/td/table/tr")
#  p rows.collect
  @long_details[sym] = rows.collect do |row|
    detail = {}
    [
      [:date, 'td[1]/text()'],
      [:close, 'td[5]/text()']
    ].each do |name, xpath|
      detail[name] = row.at_xpath(xpath).to_s.strip
    end
    detail
  end
  last5=[]
  @long_details[sym].each_with_index do |x,i|
    next if x[:close] == ""
    last5 << x[:close]
    break if last5.size == 5
  end
  sum=0
  last5.each do |x|
    sum=sum+x.to_f
  end
  avg = sum/5
  p "Symbol: #{sym}"
  p " Moving Average: #{avg}"
  p " Last Close: #{last5[0]}"
  buy = last5[0].to_f > avg
  p " Enter Long Position? #{buy}"
end

@short_details={}
shorts.keys.each do |sym|
  hp_url="http://finance.yahoo.com/q/hp?s=#{sym}+Historical+Prices"
  page = Nokogiri::HTML(open(hp_url))
  rows = page.xpath("//table[contains(@class,'yfnc_datamodoutline1')]/tr/td/table/tr")
#  p rows.collect
  @short_details[sym] = rows.collect do |row|
    detail = {}
    [
      [:date, 'td[1]/text()'],
      [:close, 'td[5]/text()']
    ].each do |name, xpath|
      detail[name] = row.at_xpath(xpath).to_s.strip
    end
    detail
  end
  last5=[]
  @short_details[sym].each_with_index do |x,i|
    next if x[:close] == ""
    last5 << x[:close]
    break if last5.size == 5
  end
  sum=0
  last5.each do |x|
    sum=sum+x.to_f
  end
  avg = sum/5
  p "Symbol #{sym}"
  p " Moving Average: #{avg}"
  p " Last close: #{last5[0]}"
  buy = last5[0].to_f < avg
  p " Buy Put? #{buy}"
end
#p @details



