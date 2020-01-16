require 'spreadsheet'

def check_digit(num)
  tot = num.digits.reverse.map.with_index do |d, idx|
    v = (((1+idx)%2)+1)
    #puts d*v
    (d * v).digits.reduce(:+)
  end.reduce(:+)
  rtot = (tot/10.to_f).ceil()*10
  #puts rtot
  #puts tot
  rtot-tot
end

if ARGV.length() != 3
  puts "USAGE: [output file] [begin range] [end range]"
  return
end

output_file = ARGV[0]
book = Spreadsheet::Workbook.new
output = []
begin_range = ARGV[1].to_i
end_range = ARGV[2].to_i

if(end_range < begin_range)
  end_range = begin_range + end_range
  puts "setting end range to "+end_range.to_s
end

(begin_range..end_range).each_slice(10000).with_index do |nums,indx|
  sheet = book.create_worksheet
  nums.each.with_index do |num,indx|
    row = sheet.row(indx)
    row.push(num.to_s+check_digit(num).to_s)
  end
end

book.write output_file
puts "XLS output to "+ARGV[0]
