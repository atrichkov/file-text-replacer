#!/usr/bin/env ruby

require 'date'
require 'docx'

class Replacer
  def initialize(dataDir)
    @dataDir = dataDir
    @source_file_path
    @destination_file_path
  end

  def run
    puts 'Enter source file name (if you press enter to skip input.docx would be used)'
    source_file = gets.chomp

    if source_file.empty?
      source_file = 'input.docx'
    end

    @source_file_path = "#{DATADIR}#{source_file}"
    @destination_file_path = "#{DATADIR}result-#{Time.now.strftime("%d-%m-%Y-%H-%M")}.docx"
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

    processDocument(replace_values)
  end

  def processDocument(replace_values)
    # Create a Docx::Document object for our existing docx file
    doc = Docx::Document.open(@source_file_path)

    doc.tables.each do |table|
      table.rows.each do |(row, val), idx| # Row-based iteration
        cell_counter = 0;
        row.cells.each do |cell|
          cell.paragraphs.each do |paragraph|
            paragraph.each_text_run do |text|
              cell_text_symbol = text.to_s.strip.to_sym
              if cell_counter == 1 and replace_values.include? cell_text_symbol
                text.substitute(text.to_s, replace_values[cell_text_symbol])
              end
            end
          end
          cell_counter += 1
        end
        cell_counter = 0
      end
    end

    doc.save(@destination_file_path)
    puts "New file #{@destination_file_path} is generated"
  end
end

DATADIR = File.join(__dir__, '../', 'data/')
replacer = Replacer.new(DATADIR)
replacer.run()
