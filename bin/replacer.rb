#!/usr/bin/env ruby

require 'docx'

DATADIR = File.join(__dir__, '../', 'data/')

puts 'Enter source file name (if you press enter to skip input.docx would be used)'
source_file = gets.chomp

if source_file == ''
  source_file = 'input.docx'
end

$source_file_path = "#{DATADIR}#{source_file}"
$destination_file_path = "#{DATADIR}result.docx"
puts 'Please enter values for replacement separated by comma (current -> new):'
replace_values = {}
input = gets.chomp
names = input.split(',')

names.each_with_index do |name, index|
  name = name.upcase.strip
  if index % 2 == 0
    replace_values[name.to_sym] = ''
  else 
    replace_values[replace_values.keys.last] = name
  end
end

puts 'Check current/new pairs for replacement and press enter to continue'
puts replace_values
gets.chomp # wait for confirmation

def processDocument(replace_values)
  # Create a Docx::Document object for our existing docx file
  doc = Docx::Document.open($source_file_path)

  doc.tables.each do |table|
    table.rows.each do |(row, val), idx| # Row-based iteration
      cell_counter = 0;
      row.cells.each do |cell|
        cell.paragraphs.each do |paragraph|
          paragraph.each_text_run do |text|
            if cell_counter == 1 and replace_values.include? paragraph.text.to_sym
              text.substitute(text.to_s, replace_values[paragraph.text.to_sym])
            end
          end
        end
        cell_counter += 1
      end
      cell_counter = 0
    end
  end

  doc.save($destination_file_path)
  puts "New file #{$destination_file_path} is generated"
end

processDocument(replace_values)
