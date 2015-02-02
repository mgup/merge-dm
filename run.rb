require 'csv'

def generate_doc_hash(row)
  student_id = row[39]
  return nil if student_id.nil?
  
  # return nil if row[34].nil?
  # diploma_number = row[34].split(' ')[1]
  # return nil if diploma_number.nil?
  
  {
    student_id: row[39],
    series: '107705' #,
    # number: diploma_number
  }
end

docs = {}

index = 1
Dir.glob('dbs/*.mdb') do |filename|
  `mdb-export -HQ -d $ -R @ #{filename} dmDoc | tr '\r' ' ' | tr '\n' ' ' > dmDoc.csv`
  
  # j = 1
  CSV.foreach('dmDoc.csv', col_sep: '$', row_sep: '@', quote_char: '§') do |row|
    student_id = row[39]
    
    hash = generate_doc_hash(row)
    
    unless hash.nil?
      if docs.has_key?(student_id)
        docs[student_id].push(generate_doc_hash(row))
      else
        docs[student_id] = [generate_doc_hash(row)]
      end
    end
              
    # j += 1
    # fail docs.inspect if 3 == j
  end
  `rm dmDoc.csv`
  
  `mdb-export -HQ -d $ -R @ #{filename} dmStudent | tr '\r' ' ' | tr '\n' ' ' > dmStudent.csv`
  CSV.foreach('dmStudent.csv', col_sep: '$', row_sep: '@', quote_char: '§') do |row|
    # puts row[39]
    # fail '123'
  end
  `rm dmStudent.csv`
  
  index += 1
  # break
end

puts docs
puts docs.size
