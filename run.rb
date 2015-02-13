require 'csv'
# require 'writeexcel'

docs = {}
reg_numbers = []

# Парсинг файлов MDB.
Dir.glob('dbs/*.mdb') do |filename|  
  # j = 1
  `mdb-export -HQ -d $ -R @ #{filename} dmDoc | tr '\r' ' ' | tr '\n' ' ' > dmDoc.csv`
  CSV.foreach('dmDoc.csv', col_sep: '$', row_sep: '@', quote_char: '§') do |row|
    reg_number = row[33].to_i
    
    student_id = row[39].to_i
    unless student_id.nil? || reg_number.nil? || [0,107705].include?(reg_number)
      reg_numbers.push(reg_number)
      
      # if docs.has_key?(reg_number)
      #   docs[reg_number][:doc].push(row)
      # else
      #   docs[reg_number] = {
      #     doc: [row],
      #     student: []
      #   }
      # end

      if docs.has_key?(student_id)
        docs[student_id][:doc].push(row)
      else
        docs[student_id] = {
          doc: [row],
          student: []
        }
      end
    end
              
    # j += 1
    # fail docs.inspect if 3 == j
  end
  `rm dmDoc.csv`
end

# fail '123'
  
Dir.glob('dbs/*.mdb') do |filename|
  puts filename
  `mdb-export -HQ -d $ -R @ #{filename} dmStudent | tr '\r' ' ' | tr '\n' ' ' > dmStudent.csv`
  CSV.foreach('dmStudent.csv', col_sep: '$', row_sep: '@', quote_char: '§') do |row|
    student_id = row[0].to_i
    if docs.has_key?(student_id)
      docs[student_id][:student].push(row)
    end
  end
  `rm dmStudent.csv`
end

# puts docs.inspect

puts "Всего импортировано записей: #{docs.size}"
puts "Уникальных регистрационных номеров: #{reg_numbers.uniq.size}"

numbers_missed = []
(1..1343).each do |n|
  unless reg_numbers.include?(n)
    numbers_missed.push(n)
  end
end

if numbers_missed.any?
  puts "Не хватает данных для регистрационных номеров (#{numbers_missed.size}): #{numbers_missed.sort.join(', ')}."
end

clean_docs = []

docs.each do |reg_number, doc|
  # ФИО.
  values = []
  doc[:doc].each { |row| values.push(row[28]) unless row[28].nil? || '' == row[28] }
  if values.uniq.size != 1
    r28 = [nil, nil, nil]
  else
    # puts values.uniq.first.strip.split(' ').inspect
    if values.uniq.first.strip.split(' ').size != 3
      r28 = [nil, nil, nil]
    else
      r28 = values.uniq.first.strip.split(' ')
    end
  end

  # Дата рождения.
  values = []
  doc[:doc].each do |row| 
    unless row[7].nil? || '' == row[7]
      begin
        parsed = Date.strptime(row[7], '%m/%d/%y 00:00:00')
      rescue
        fail row[7].inspect
      end
      
      # unless Date.parse('05/05/2014') == parsed
      unless 2014 == parsed.year
        values.push(parsed)
      end
    end
  end
  if values.uniq.size != 1
    r7 = nil
  else
    r7 = values.uniq.first
    # puts r7
  end
  
  # Пол.
  values = []
  doc[:doc].each { |row| values.push(row[8]) unless row[8].nil? || '' == row[8] }
  if values.uniq.size != 1
    r8 = nil
  else
    r8 = values.uniq.first
  end
  
  # Специальность.
  type_bak = false
  type_mag = false
  type_spec = false
  values = []
  doc[:doc].each { |row| values.push(row[32]) unless row[32].nil? || '' == row[32] }
  if values.uniq.size != 1
    r32 = ['', '']
  else
    r32 = ['', '']
    val = values.uniq.first
    r32[0] = val[0..(val.index(' ') - 1)]
    r32[1] = val[(val.index(' ') + 1)..val.size]
    
    type_bak = true if '62' == r32[0].split('.')[1]
    type_mag = true if '68' == r32[0].split('.')[1]
    type_spec = true if '65' == r32[0].split('.')[1]
  end
  
  # Квалификация.
  values = []
  doc[:doc].each { |row| values.push(row[30]) unless row[30].nil? || '' == row[30] }
  if values.uniq.size != 1
    r30 = ''
  else
    r30 = values.uniq.first
  end
  
  # С отличием
  dip_cool = false
  values = []
  doc[:doc].each { |row| values.push(row[10]) unless row[10].nil? || '' == row[10] }
  unless values.uniq.size != 1
    dip_cool = true if '1' == values.uniq.first
  end
  
  # Вид документа
  doc_type = ''
  edu_type = ''
  if type_bak
    edu_type = 'Высшее образование - бакалавриат'
    if dip_cool
      doc_type = 'Диплом бакалавра с отличием'
    else
      doc_type = 'Диплом бакалавра'
    end
  end
  if type_mag
    edu_type = 'Высшее образование - специалитет, магистратура'
    if dip_cool
      doc_type = 'Диплом магистра с отличием'
    else
      doc_type = 'Диплом магистра'
    end
  end
  if type_spec
    edu_type = 'Высшее образование - специалитет, магистратура'
    if dip_cool
      doc_type = 'Диплом специалиста с отличием'
    else
      doc_type = 'Диплом специалиста'
    end
  end
  
  # Номер диплома.
  values = []
  doc[:doc].each { |row| values.push(row[34]) unless row[34].nil? || '' == row[34] }
  if values.uniq.size != 1
    r34 = ''
  else
    if values.uniq.first.split(' ').size != 2
      r34 = ['', '']
    else
      r34 = values.uniq.first.split(' ')
    end
  end  
  
  # Регистрационный номер.
  values = []
  doc[:doc].each { |row| values.push(row[33]) unless row[33].nil? || '' == row[33] }
  if values.uniq.size != 1
    r33 = ''
  else
    r33 = values.uniq.first
  end
  
  # Дата выдачи
  values = []
  doc[:doc].each do |row| 
    unless row[5].nil? || '' == row[5]
      begin
        parsed = Date.strptime(row[5], '%m/%d/%y %H:%M:%S')
      rescue
        fail row[5].inspect
      end
      
      # unless Date.parse('05/05/2014') == parsed
      # unless 2014 == parsed.year
      #   values.push(parsed)
      # end
      values.push(parsed)
    end
  end
  if values.uniq.size != 1
    r5 = nil
  else
    r5 = values.uniq.first
    # puts r7
  end
  
  clean_docs.push({
    'Название документа'  => 'Диплом',
    'Вид документа'  => doc_type,
    'Статус документа'  => nil,
    'Подтверждение утраты'  => nil,
    'Подтверждение обмена'  => nil,
    'Уровень образования'  => edu_type,
    'Серия документа'  => r34[0],
    'Номер документа'  => r34[1],
    'Дата выдачи'  => r5.nil? ? nil : r5.strftime('%d.%m.%Y'),
    'Регистрационный номер'  => r33,
    'Код специальности, направления подготовки'  => r32[0],
    'Наименование специальности, направления подготовки'  => r32[1],
    'Наименование квалификации'  => r30,
    'Образовательная программа'  => '',
    'Год поступления'  => '',
    'Год окончания'  => '2014',
    'Срок обучения, лет'  => '',
    'Фамилия получателя'  => r28[0],
    'Имя получателя'  => r28[1],
    'Отчество получателя'  => r28[2],
    'Дата рождения получателя'  => r7.nil? ? nil : r7.strftime('%d.%m.%Y'),
    'Пол получателя'  => r8.nil? ? nil : ('1' == r8 ? 'Муж' : 'Жен')
  })
end

# puts reg_numbers.uniq.min
# puts reg_numbers.uniq.max

# Начинаем анализ данных. Удаляем всех студентов, у которых не заполнено поле РегНомер
# to_delete = []
# docs.each do |student_id, data|
#   has_dp_name = false
#   data[:student].each do |data_row|
#     if 'Чумаков' == data_row[1] || !data_row[4].nil?
#       puts data_row.inspect
#     end
#
#     has_dp_name ||= !data_row[4].nil?
#   end
#
#   to_delete.push(student_id) unless has_dp_name
# end
# to_delete.each { |key| docs.delete(key) }

puts clean_docs[0..2].inspect

CSV.open('results.csv', 'wb') do |csv|
  clean_docs.each do |hash|
    csv << hash.values
  end
end