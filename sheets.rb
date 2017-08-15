module Sheets

def self.fill_out_sheet(sheet, transactions)

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
      t['Account ID'],
      t['Paycode Name'],
      t['Account Name'],
      t['First Name'],
      t['Last Name'],
      t['Home Phone'],
      t['Mobile Phone'],
      t['All Email Addresses'],
      t['Transaction Type'],
      t['Informal Salutation']
  end
end
