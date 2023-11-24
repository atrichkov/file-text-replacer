#!/usr/bin/env ruby

require 'docx'

DATADIR = File.join(__dir__, '../', 'data/')

puts 'Enter source file name (if you press enter to skip input.docx would be used)'
sourceFile = gets.chomp

if sourceFile == ""
  sourceFile = 'input.docx'
end

$sourceFileName = "#{DATADIR}#{sourceFile}"
puts $sourceFileName
$destinationFileName = "#{DATADIR}result.docx"
puts "Please enter values for replacement separated by comma (current -> new):"
ReplaceValues = {}
input = gets.chomp
names = input.split(',')

names.each_with_index do |name, index|
  name = name.upcase.strip
  if index % 2 == 0
    ReplaceValues[name.to_sym] = ""
  else 
    ReplaceValues[ReplaceValues.keys.last] = name
  end
end

puts "Check current/new pairs for replacement and press enter to continue"
puts ReplaceValues
gets.chomp # wait for confirmation

def processDocument(replaceValues)
  # Create a Docx::Document object for our existing docx file
  doc = Docx::Document.open($sourceFileName)

  doc.tables.each do |table|
    table.rows.each do |(row, val), idx| # Row-based iteration
      counter = 0;
      row.cells.each do |cell|
        cell.paragraphs.each do |paragraph|
          paragraph.each_text_run do |text|
            if counter == 1 and replaceValues.include? paragraph.text.to_sym
              text.substitute(text.to_s, replaceValues[paragraph.text.to_sym])
            end
          end
        end
        counter += 1
      end
      counter = 0
    end
  end

  doc.save($destinationFileName)
  puts "New file #{$destinationFileName} is generated"
end

processDocument(ReplaceValues)
