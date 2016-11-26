
require 'rspreadsheet'
require 'byebug'

book = Rspreadsheet.open('./min.ods')
sheet = book.worksheets(1)

(1..2000).each do |n|
  puts n
  sheet[8,2] = n
end

book.save('min2.ods')
