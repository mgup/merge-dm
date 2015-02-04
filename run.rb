require 'csv'

docs = {}

Dir.glob('dbs/*.mdb') do |filename|
  `mdb-export -HQ -d $ -R @ #{filename} dmDoc | tr '\r' ' ' | tr '\n' ' ' > dmDoc.csv`
  
  # j = 1
  CSV.foreach('dmDoc.csv', col_sep: '$', row_sep: '@', quote_char: '§') do |row|
    student_id = row[39]
    unless student_id.nil?
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
  
  `mdb-export -HQ -d $ -R @ #{filename} dmStudent | tr '\r' ' ' | tr '\n' ' ' > dmStudent.csv`
  CSV.foreach('dmStudent.csv', col_sep: '$', row_sep: '@', quote_char: '§') do |row|
    student_id = row[0]
    if docs.has_key?(student_id)
      docs[student_id][:student].push(row)
    end
  end
  `rm dmStudent.csv`
  
  # break
end

puts docs
puts docs.size
