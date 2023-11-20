#!/usr/bin/env ruby

require 'docx'

DATADIR = File.join(__dir__, '../', 'data/')
$sourceFileName = "#{DATADIR}input.docx";
$destinationFileName = "#{DATADIR}result.docx";
puts "Please enter values for replacement separated by spaces (current -> new):"
ReplaceValues = {}
input = gets.chomp
names = input.split

names.each_with_index do |name, index|
  next if name == 'done'
  if index % 2 == 0
    ReplaceValues[name.upcase.to_sym] = ""
  else 
    ReplaceValues[ReplaceValues.keys.last] = name.upcase
  end
end

puts "Check current/new paris for replacement and press enter to continue"
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
