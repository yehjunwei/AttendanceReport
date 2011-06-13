# coding: utf-8
#!/usr/bin/env ruby 1.9.2

require 'rubygems'
require 'roo'
require 'date'

pwd       = File.dirname(__FILE__)

class ReportGenerator

	def initialize(year, month, intern, late_hour, late_min)
		@target_year = year
		@target_month = month
		@intern = intern
		@late_hour = late_hour
		@late_min = late_min
	end
	
	def generate(xls, output_csv)
		number = xls.cell(2,'A')
		name = xls.cell(2,'B')
		
		# Decide if this employee is intern
		@is_intern = false
		if @intern.include?(xls.cell(2,'A').to_i)
			@is_intern = true
		end
		
		last_date_time = nil
		start_time = nil
		end_time = nil
		
		xls.last_row.downto(2) do |line|
			date_time_s = xls.cell(line,'D').to_s.split
			# date
			date_s = date_time_s[0]
			date = Date.parse(date_s)
			
			# if year and month are not match, continue loop
			next if date.month != @target_month || date.year != @target_year
			
			# time
			time_s = date_time_s[1]
			time = Time.parse(time_s)
			
			# create DateTime object according to the current time
			date_time = DateTime.new(date.year, date.month, date.day, time.hour, time.min, time.sec)
			
			# handle the first day of month, like 5/1, make last_day = 0
			current_day = date_time.day
			if last_date_time == nil
				last_day = 0
			else
				last_day = last_date_time.day
			end
			
			# calculate working hours when the day increased
			if current_day != last_day
				if last_day != 0
					# store end_time, calculate!
					end_time = last_date_time
					calculate_working_hours(number, name, start_time, end_time, output_csv)
				end
				# store start_time
				end_time = nil
				start_time = date_time
			end
			
			# calculate the calculate_working_hours on the last day like 5/31
			if line == 2
				end_time = date_time
				calculate_working_hours(number, name, start_time, end_time, output_csv)
			end
			
			last_date_time = date_time
		end
	end
	
	private
	def calculate_working_hours(number, name, start_time, end_time, output_csv)
		record = {"late" => false, "no_record" => false, "not_enough_time" => false}
		
		if start_time == end_time
			# only 1-record detection
			#puts "#{start_time}\tOnly 1 record on this day"
			record["no_record"] = true
		else
			# late start detection
			if start_time.hour > @late_hour
				#puts "#{start_time}\tLate"
				record["late"] = true
			elsif start_time.hour == @late_hour
				if start_time.min > @late_min
					#puts "#{start_time}\tLate"
					record["late"] = true
				end
			end

			# not enough working hour detection
			working_hour = ((end_time.hour * 60 + end_time.min) - (start_time.hour * 60 + start_time.min)) / 60.0
			
			if working_hour < 9.0
				#puts "#{end_time}\tOnly #{working_hour.round(2)} hours"
				record["not_enough_time"] = true
			end
		end
		
		# print record
		if(record["late"] || record["no_record"] || record["not_enough_time"])
			output_csv.syswrite("#{number},#{name},#{start_time.strftime("%Y-%m-%d %H:%M:%S")},#{end_time.strftime("%Y-%m-%d %H:%M:%S")},#{record["late"]? 1:nil},#{record["no_record"]? 1:nil},")
			if(record["not_enough_time"])
				output_csv.syswrite("#{record["not_enough_time"]? 1:nil},#{9.0-working_hour.round(2)}\n")
			else
				output_csv.syswrite(", \n")
			end
		end
	end

end

#define target year/month, intern employee number, late hour/minute
target_year = 2011
target_month = 6
intern = [22,23,24]
late_hour = 9
late_min = 40

doorman = ReportGenerator.new(target_year, target_month, intern, late_hour, late_min)

output_csv = File.new("#{pwd}/#{target_year}_#{target_month}_Report.csv", "w+")
if output_csv
	output_csv.syswrite("Number,Name,start,end,late,no record,not enough time,lacking hours\n")
else
	puts "Unable to open file!"
end

Dir.glob("#{pwd}/*.xls") do |file|
	file_path = "#{pwd}/#{file}"  
	file_basename = File.basename(file, ".xls")  
	xls = Excel.new(file_path)
	doorman.generate(xls, output_csv)
end

output_csv.close
