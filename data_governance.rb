require 'creek'
require 'csv'

class DataGvernance
  def perform
    creek = Creek::Book.new '/Users/mac/Desktop/工作/合作/一牌一薄/合作辖区从业人员.xlsx'
    sheet = creek.sheets[0]
    need_data(sheet)
  end

  def need_data(sheet)
    sheet.simple_rows.each_with_object([]) do |row, unit_name|
      csv_file(row).write([csv_headers].inject([]) { |csv, row| csv << CSV.generate_line(row) }.join('')) if (unit_name & [row['F']]).empty?
      unit_name << row['F']
      next if row['D'] == '负责人' or row['F'] == '单位名称'
      puts "current unit: #{row['F']}"
      csv_file(row).write([current_rows(row)].inject([]) { |csv, csv_row| csv << CSV.generate_line(csv_row) }.join(''))
      #csv_file(row) << current_rows(row)
    end
  end

  def csv_headers
    %w(姓名 出生日期 证件号码 职务类别 联系电话 单位名称 录入时间 单位分类 创建人 所属派出所 性别 国籍 籍贯 可疑迹象 可疑行为 最后修改时间 注销标志 是否有联系电话 所在部门 所在岗位 开始日期 结束日期 单位地址 民族 是否QBZDR 是否ZAZDR)
  end

  def current_rows(row)
    [row['A'], row['B'], row['C'], row['D'], row['E'], row['F'], row['G'], row['H'], row['I'], row['J'], row['K'], row['L'],
     row['M'], row['N'], row['O'], row['P'], row['Q'], row['R'], row['S'], row['T'], row['U'], row['V'], row['W'], row['X'],
     row['Y'], row['Z']]
  end

  def csv_file(row)
    File.open(current_file_name(row), 'a+')
  end

  def current_file_name(row)
    "./employee/#{row['F']}.csv"
  end
end

DataGvernance.new.perform
