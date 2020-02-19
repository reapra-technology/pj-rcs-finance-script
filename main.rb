#!/usr/bin/ruby

require 'optparse'
require 'optparse/time'
require 'ostruct'
require 'pp'
require 'rubyXL'
require 'roo'
require 'pry'
require 'pry-byebug'

# Begin autoloading
require 'zeitwerk'
require 'listen'
loader = Zeitwerk::Loader.new
loader.enable_reloading

loader.push_dir("lib")
loader.push_dir("script")

loader.setup

Listen.to("lib", "script") { loader.reload }.start
loader.eager_load
# End autoloading

class OptParser

  CODES = %w[iso-2022-jp shift_jis euc-jp utf8 binary]
  CODE_ALIASES = { "jis" => "iso-2022-jp", "sjis" => "shift_jis" }

  #
  # Return a structure describing the options.
  #
  def self.parse(args)
    # The options specified on the command line will be collected in *options*.
    # We set default values here.
    options = OpenStruct.new
    options.input = nil
    options.output = nil
    options.bs = nil
    options.is = nil
    options.cje = nil
    options.process = []
    options.verbose = false

    opt_parser = OptionParser.new do |opts|
      opts.banner = "Usage: main.rb [options]"

      opts.separator ""
      opts.separator "Specific options:"

      # Mandatory argument.
      opts.on("-i", "--input EXCELX",
              "The xlsx file which is export from Netsuite") do |input|
        options.input = input
      end

      opts.on("-o", "--output EXCELX",
              "The xlsx file to save after processed") do |output|
        options.output = output || "./output/output.xlsx"
      end

      # Optional argument; multi-line description.
      opts.on("--balancesheet=STRING",
              "The balance sheet you need to proceed to generate consolidation report",
              "  (optional)") do |bs|
        options.bs = bs
        # options.extension.sub!(/\A\.?(?=.)/, ".")  # Ensure extension begins with dot.
      end

      opts.on("--incomestatement=STRING",
              "The income statement you need to proceed to generate consolidation report",
              "  (optional)") do |is|
        options.is = is
        # options.extension.sub!(/\A\.?(?=.)/, ".")  # Ensure extension begins with dot.
      end

      opts.on("--cje=STRING",
              "The CJE sheet for referencing in Consolidation BS/IC",
              "  default: \"CJE\" (optional)") do |cje|
        options.cje = cje || "CJE"
        # options.extension.sub!(/\A\.?(?=.)/, ".")  # Ensure extension begins with dot.
      end

      # List of arguments.
      opts.on("-p", "--process cje,bs,is", Array, "To let script execute determined tasks",
        "  cje: generate CJE sheet (which is mandatory for BS/IS process)",
        "   bs: generate Consolidation_BS sheet",
        "   is: generate Consolidation_IS sheet") do |process|
        options.process = process
      end

      # Boolean switch.
      opts.on("-v", "--[no-]verbose", "Run verbosely with log") do |v|
        options.verbose = v
      end

      opts.separator ""
      opts.separator "Common options:"

      # No argument, shows at tail.  This will print an options summary.
      # Try it and see!
      opts.on_tail("-h", "--help", "Show this message") do
        puts opts
        exit
      end

      # Another typical switch to print the version.
      opts.on_tail("--version", "Show semver") do
        puts ["0","9","1"].join('.')
        exit
      end

      opts.on_tail("--example", "Show example command line"
        ) do
        puts "---- TO GENERATE CJE ----"
        puts "bundle exec ruby main.rb -i 'input/consolidation.xlsx' -o 'output/new_consolidation.xlsx' -p cje --cje=CJE"
        puts
        puts "---- TO GENERATE CONSOLIDATION REPORTS ----"
        puts "bundle exec ruby main.rb -i 'input/consolidation.xlsx' -o 'output/new_consolidation.xlsx' -p bs,is --balancesheet='BalanceSheet 15-8-19' --incomestatement='IncomeStatement 15-8-19' --cje=CJE"
        puts
        puts "---- TO GENERATE CONSOLIDATION BALANCE SHEET WITH ERRORS / PROCESSING STATUS ----"
        puts "bundle exec ruby main.rb -i 'input/consolidation.xlsx' -o 'output/new_consolidation.xlsx' -p bs --balancesheet='BalanceSheet 15-8-19' --cje=CJE -v"

        exit
      end
    end

    opt_parser.parse!(args)
    options
  end  # parse()

end  # class OptParser

options = OptParser.parse(ARGV)
pp options

# check required arguments
raise OptionParser::MissingArgument.new("input") if options[:input].nil?
raise OptionParser::MissingArgument.new("bs") if options[:bs].nil? && options[:process].include?("bs")
raise OptionParser::MissingArgument.new("is") if options[:is].nil? && options[:process].include?("is")
raise OptionParser::MissingArgument.new("process cje require both balancesheet and incomestatement to be listed") if (options[:is].nil? || options[:bs].nil?) && options[:process].include?("cje")
raise OptionParser::MissingArgument.new("no process defined to run") if options[:process].empty?

# FIXME refactor if below to be switch case
if options[:process].include? "cje"
  NetsuiteConsolidationReportCjeBuilder.new(options[:input], options[:bs], options[:is], options[:cje], options[:v]).build_cje_sheet

  return
end

if options[:process].include?("bs") || options[:p].include?("is")
  ns = NetsuiteConsolidationReport.new(options[:input], options[:output], options[:bs], options[:is], options[:cje], options[:v]).prerequisite
  ns.clone_entities_with_ref_formulas
  ns.run

  return
end

