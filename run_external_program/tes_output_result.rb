# coding: Windows-31J

=begin
** coding ===>>>
shift_jis
ASCII-8BIT
=end

require './common'

#p "テスト"
#p "テスト".encoding

if ARGV.size != 2 then
  output_result(Dir.pwd + "/result.txt", -1, "argument error!!\n")
  exit
end

bn = ARGV[0]
sn = ARGV[1]

ao = WIN32OLE.new('Excel.Application')

begin
  code = -2
  bo = ao.Workbooks.Open(Dir.pwd + "/" + bn)
  code = -3
  ary = read_sheet_data(bo.Worksheets(sn), 2, 10, 1, 2)
  code = -4
  open(Dir.pwd + "/tes_output_result.txt", "w") do |f|
    ary.each_with_index do |a1, i|
      a1.each_with_index do |a2, j|
        f.print "ary[" + i.to_s + ", " + j.to_s + "]: " + ary[i][j] + "\n"
      end
    end
  end
  output_result(Dir.pwd + "/result.txt", 0, "normal end!!!\n")
rescue
  output_result(Dir.pwd + "/result.txt", code, $@[1] + ": " + $!.message + "\n")
#  msg = $!.message.gsub(/\n/, " ")
#  p $!.message.encoding
#  p msg.encoding
#  p msg.encode("ASCII-8BIT")
#  p msg.encode("Windows-31J")
#  p msg.encode("shift_jis")
#  p $!.message.encode("ASCII-8BIT")
#  p $!.message.encode("UTF-8")
ensure
  bo.Close if defined?(bo) and not bo.nil?
  ao.Quit if defined?(ao) and not ao.nil?
end
