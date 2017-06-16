module Sheets

def self.fill_out_sheet(sheet, transactions)

  transactions.each_with_index do |t, i|
    sheet.update_row (i+1),
      t['Transaction Date'],
      t['User Defined Field 2'],
      t['Transaction Amount'],
      t['Receipted Account Name'],
      t['Fund #'],
      t['Pay Method'],
      t['Address Line 1'],
      t['Address Line 2'],
      t['Address Line 3'],
      t['City'],
      t['State'],
      t['Zip Code'],
      t['Allocation Name'],
      t['AD Number'],
      t['Payment Type'],
      t['Account Name'],
      t['First Name'],
      t['Last Name'],
      t['Home Phone'],
      t['Mobile Phone'],
      t['Work Phone'],
      t['All Email Addresses'],
      t['Transaction Type'],
      t['Salutation'],
      t['Attention Name'],
      t['Company'],
      t['County'],
      t['Country']
  end
end

def self.fill_out_xos(sheet, transactions)

  transactions.each_with_index do |t, i|
    sheet.update_row (i+1), t['Transaction Date'],
             t['AD Number'],
             t['User Defined Field 2'],
             t['Address Line 1'],
             t['Address Line 2'],
             t['Address Line 3'],
             t['City'],
             t['State'],
             t['Zip Code'],
             t['Transaction Amount'],
             t['Receipted Account Name'],
             t['Fund #'],
             t['Pay Method'],
             t['Allocation Name'],
             t['Payment Type'],
             t['Account Name'],
             t['First Name'],
             t['Last Name'],
             t['Home Phone'],
             t['Mobile Phone'],
             t['Work Phone'],
             t['All Email Addresses'],
             t['Transaction Type'],
             t['Salutation'],
             t['Attention Name'],
             t['Company'],
             t['County'],
             t['Country']
    end
  end
end
