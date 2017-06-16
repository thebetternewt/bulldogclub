require './sheets.rb'
require 'fileutils'
require 'csv'
require 'spreadsheet'

bdc_file = ARGV.first

# Cleanup old .xls files in directory.
FileUtils.rm Dir.glob('*.xls')

# Read funds from csv.
funds = CSV.read('fund_dictionary.csv', headers: true)

fund_names = funds['Fund Name'] # Extract Fund Names to array.
fund_numbers = funds['Fund Number'] # Extract Fund Numbers to array.
funds = fund_names.zip(fund_numbers).to_h # Convert funds to hash.

# Read gifts from csv.
gift_headers = CSV.read(bdc_file, headers: true).headers
# gifts = CSV.read('bdc_orig.csv', headers: true)

gifts = []
CSV.foreach(bdc_file, headers: true) do |gift|
  gifts << gift
end

# Create Excel workbook with sheets.
book = Spreadsheet::Workbook.new
original = book.create_worksheet name: 'ORIGINAL'
entry = book.create_worksheet name: 'ENTRY'
adj = book.create_worksheet name: 'ADJ'
data_mgt = book.create_worksheet name: 'DATA MGT'
xos = book.create_worksheet name: 'XOS'

# Add ORIGINAL headers.
gift_headers.each do |header|
  original.row(0).push header
end

# Add ORIGINAL records.
gifts.each_with_index do |gift, i|
  # Add each attribute's value to the row (i).
  gift.each { |e| original.row(i + 1).push e.last }
end

# =================
# Clean up gifts.
# =================

# Clean Banner IDs ('User Defined Field 2' on BC Report).
gifts.each do |g|
  # Convert nil Banner IDs to empty strings.
  g['User Defined Field 2'] = g['User Defined Field 2'].to_s
  # Remove any dashes.
  g['User Defined Field 2'].gsub!(/[-]/, '')
  # Strip any extra numbers.
  g['User Defined Field 2'] = g['User Defined Field 2'][0..8]
  # Set invalid IDs to blank strings.
  unless g['User Defined Field 2'] =~ /^[0-9]{9}$/
    g['User Defined Field 2'] = ''
  end
end

# Clean phone numbers.
gifts.each do |g|
  [g['Home Phone'], g['Mobile Phone'], g['Work Phone']].each do |phone|
    phone = phone.to_s
    phone.gsub!(/[-()_\.\s]/, '') # Remove any symbols
    phone.insert(3, '-') unless phone.length < 3 # Add first hyphen.
    phone.insert(7, '-') if phone.length >= 8 # Add second hyphen.
  end
end

# Sort gifts by Banner ID.
gifts.sort_by! { |gift| gift['User Defined Field 2'] }

# Calculate gift_total
gift_total = 0
gifts.each { |g| gift_total += g['Transaction Amount'].to_f }
gift_total = gift_total.round(2)

# Update gifts.
gifts.each do |g|
  # Update Fund Number.
  if !funds[g['Allocation Name']].nil?
    g['Fund #'] = funds[g['Allocation Name']]
  else
    g['Fund #'] = '000000'
  end

  # Update Pay Method.
  if g['Payment Type'] == 'GIK'
    g['Pay Method'] = 'BI'
  else
    g['Pay Method'] = 'BC'
  end
end

# ==================
# Find adjustments.
# ==================

# Find adjustment_ids (AD Numbers).
adjustment_ids = []
gifts.each do |g|
  adjustment_ids << g['AD Number'] if g['Transaction Amount'].to_f < 0
end

# Fund adjustments from adjustment_ids.
adjustments = []
gifts.each { |g| adjustments << g if adjustment_ids.include?(g['AD Number']) }

# Remove adjustments from gifts list.
gifts -= adjustments

# ==============================
# Find gifts with no Banner IDs.
# ==============================

gifts_with_no_id = []
gifts.each do |gift|
  gifts_with_no_id << gift if gift['User Defined Field 2'].empty?
end

entry_headers = 'Paid Date',
                'Banner ID',
                'Amount Paid',
                'Receipted Account Name',
                'Fund #',
                'Pay Method',
                'Address Line 1',
                'Address Line 2',
                'Address Line 3',
                'City',
                'State',
                'Zip Code',
                'Fund Name',
                'XOS Acct #',
                'Transaction Type',
                'Account Name',
                'First Name',
                'Last Name',
                'Home Phone',
                'Mobile Phone',
                'Work Phone',
                'All Email Addresses',
                'Transaction Type',
                'Salutation',
                'Attention Name',
                'Company',
                'County',
                'Country'

xos_headers =  'Paid Date',
               'XOS Acct #',
               'Banner ID',
               'Address Line 1',
               'Address Line 2',
               'Address Line 3',
               'City',
               'State',
               'Zip Code',
               'Amount Paid',
               'Receipted Account Name',
               'Fund #',
               'Pay Method',
               'Fund Name',
               'Transaction Type',
               'Account Name',
               'First Name',
               'Last Name',
               'Home Phone',
               'Mobile Phone',
               'Work Phone',
               'All Email Addresses',
               'Transaction Type',
               'Salutation',
               'Attention Name',
               'Company',
               'County',
               'Country'

# Add ENTRY, ADJ, DATA MGT headers.
[entry, adj, data_mgt].each do |sheet|
  entry_headers.each do |header|
    sheet.row(0).push header
  end
end

# Add XOS headers.
xos_headers.each do |header|
  xos.row(0).push header
end

# Add ENTRY records.
Sheets.fill_out_sheet(entry, gifts)
Sheets.fill_out_sheet(adj, adjustments)
Sheets.fill_out_sheet(data_mgt, gifts_with_no_id)
Sheets.fill_out_xos(xos, gifts)

# Generate Excel spreadsheet.
wb_name = "bdc_report_#{DateTime.now.strftime('%y%m%dT%H%M%S%z')}.xls"
book.write "./#{wb_name}"

# Open file upon completion.
if /cygwin|mswin|mingw|bccwin|wince|emx/ =~ RUBY_PLATFORM # Check if Windows OS
  system %{ cmd /c "start #{wb_name}" }
else system %{ open "#{wb_name}" } # Assume Mac OS/Linux
end
