require 'fileutils'
require 'csv'
require 'spreadsheet'

def fill_out_sheet(sheet, transactions)
  transactions.each_with_index do |t, i|
    sheet.update_row (i+1),
      t['Received Date'],
      t['Banner ID'],
      t['Payment Amount'],
      t['Receipted Account Name'],
      t['Fund #'],
      t['Pay Method'],
      t['Fund Bill Address Line1'],
      t['Fund Bill Address Line2'],
      t['Fund Bill Address Line3'],
      t['Fund Bill City'],
      t['Fund Bill State'],
      t['Fund Bill Zip Code'],
      t['Allocation'],
      t['Paycode Name'],
      t['Account Name'],
      t['First Name'],
      t['Last Name'],
      t['Phone Dev Home'],
      t['Phone Dev Cell'],
      t['Informal Salutation']
  end
end

bdc_file = ARGV.first

# Cleanup old .xls files in directory.
FileUtils.rm Dir.glob('*.xls')

# Read funds from csv.
funds = CSV.read('fund_dictionary.csv', headers: true)

fund_names = funds['Fund Name'] # Extract Fund Names to array.
fund_numbers = funds['Fund Number'] # Extract Fund Numbers to array.
funds = fund_names.zip(fund_numbers).to_h # Convert funds to hash.

# Read gifts from csv.
gift_headers = CSV.read(bdc_file, headers: true, encoding: 'windows-1251:utf-8').headers
# gifts = CSV.read('bdc_orig.csv', headers: true)

gifts = []
CSV.foreach(bdc_file, headers: true, encoding: 'windows-1251:utf-8') do |gift|
  unless (gift['Allocation'] == 'Facilities Rental Facilities') ||
         (gift['Allocation'] == 'Facilities Rental Suites')
    gifts << gift
  end
end

# Create Excel workbook with sheets.
book = Spreadsheet::Workbook.new
original = book.create_worksheet name: 'ORIGINAL'
entry = book.create_worksheet name: 'ENTRY'
adj = book.create_worksheet name: 'ADJ'
data_mgt = book.create_worksheet name: 'DATA MGT'

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

# Clean Banner IDs ('Banner ID' on BC Report).
gifts.each do |g|
  # Convert nil Banner IDs to empty strings.
  g['Banner ID'] = g['Banner ID'].to_s
  # Remove any dashes.
  g['Banner ID'].gsub!(/[-]/, '')
  # Strip any extra numbers.
  g['Banner ID'] = g['Banner ID'][0..8]
  # Set invalid IDs to blank strings.
  unless g['Banner ID'] =~ /^[0-9]{9}$/
    g['Banner ID'] = ''
  end
end

# Clean phone numbers.
gifts.each do |g|
  [g['Phone Dev Home'], g['Phone Dev Cell']].each do |phone|
    phone = phone.to_s
    phone.gsub!(/[-()_\.\s]/, '') # Remove any symbols
    phone.insert(3, '-') unless phone.length < 3 # Add first hyphen.
    phone.insert(7, '-') if phone.length >= 8 # Add second hyphen.
  end
end

# Sort gifts by Banner ID.
gifts.sort_by! { |gift| [gift['Banner ID'], gift['Account ID']] }

# Calculate gift_total
gift_total = 0
gifts.each { |g| gift_total += g['Payment Amount'].to_f }
gift_total = gift_total.round(2)

# Update gifts.
gifts.each do |g|
  # Update Fund Number.
  if !funds[g['Allocation']].nil?
    g['Fund #'] = funds[g['Allocation']]
  else
    g['Fund #'] = '000000'
  end

  # Update Pay Method.
  if g['Paycode Name'] == 'GIK'
    g['Pay Method'] = 'BI'
  else
    g['Pay Method'] = 'BC'
  end
end

# ==================
# Find adjustments.
# ==================

# Find adjustment_ids (Account IDs).
adjustment_ids = []
gifts.each do |g|
  adjustment_ids << g['Account ID'] if g['Payment Amount'].to_f < 0
end

# Fund adjustments from adjustment_ids.
adjustments = []
gifts.each { |g| adjustments << g if adjustment_ids.include?(g['Account ID']) }

# Remove adjustments from gifts list.
gifts -= adjustments

# ==============================
# Find gifts with no Banner IDs.
# ==============================

gifts_with_no_id = []
gifts.each do |gift|
  gifts_with_no_id << gift if gift['Banner ID'].empty?
end

entry_headers = 'Paid Date',
                'Banner ID',
                'Amount Paid',
                'Account Name',
                'Fund #',
                'Pay Method',
                'Address Line 1',
                'Address Line 2',
                'Address Line 3',
                'City',
                'State',
                'Zip Code',
                'Fund Name',
                'Transaction Type',
                'Account Name',
                'First Name',
                'Last Name',
                'Home Phone',
                'Mobile Phone',
                'Salutation'

# Add ENTRY, ADJ, DATA MGT headers.
[entry, adj, data_mgt].each do |sheet|
  entry_headers.each do |header|
    sheet.row(0).push header
  end
end

# Add ENTRY records.
fill_out_sheet(entry, gifts)
fill_out_sheet(adj, adjustments)
fill_out_sheet(data_mgt, gifts_with_no_id)

# Generate Excel spreadsheet.
wb_name = "bdc_report_#{DateTime.now.strftime('%y%m%dT%H%M%S%z')}.xls"
book.write "./#{wb_name}"

# Open file upon completion.
if /cygwin|mswin|mingw|bccwin|wince|emx/ =~ RUBY_PLATFORM # Check if Windows OS
  system %{ cmd /c "start #{wb_name}" }
else system %{ open "#{wb_name}" } # Assume Mac OS/Linux
end
