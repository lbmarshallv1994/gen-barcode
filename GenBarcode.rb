require 'spreadsheet'

# generate the check digit
def check_digit(num)
  tot = num.digits.reverse.map.with_index do |d, idx|
    v = (((1+idx)%2)+1)
    (d * v).digits.reduce(:+)
  end.reduce(:+)
  rtot = (tot/10.to_f).ceil()*10
  rtot-tot
end

# check that user is putting in enough args
if ARGV.length() != 3
  puts "USAGE: [output file] [begin range] [end range]"
  return
end


output_file = ARGV[0]
# create spreadsheet object to dump data into
book = Spreadsheet::Workbook.new
# set the ranges, ruby args come in as strings, so we must convert them
begin_range = ARGV[1].to_i
end_range = ARGV[2].to_i

# this is what lets you create a certain number of barcodes beyond the start range
if(end_range < begin_range)
  end_range = begin_range + end_range
  puts "setting end range to "+end_range.to_s
end

# iterate over the whole range of codes, in slcies of 10,000 codes
(begin_range..end_range).each_slice(10000).with_index do |nums,indx|
  # create a sheet to hold this set of codes
  sheet = book.create_worksheet
  nums.each.with_index do |num,indx|
    # create a new row for each code
    row = sheet.row(indx)
    # push the modified code into the row
    row.push(num.to_s+check_digit(num).to_s)
  end
end

puts "Checkdigit process complete, outputting to file."

# being putting data into XLS
book.write output_file
puts "XLS output to "+ARGV[0]
