#!/usr/bin/env ruby
require 'rubygems'
require 'roo'
require 'date'

pwd       = File.dirname(__FILE__)

def calculate_working_hours(start_time, end_time)
	# define late hour/minute
	late_hour = 9
	late_minute = 40

	# if start_time == end_time means it only has one record in a day	
	if start_time == end_time
		puts "#{start_time}\tOnly 1 record on this day"
		return
	end
	
	# late start detection
	if start_time.hour > late_hour
		puts "#{start_time}\tLate"
	elsif start_time.hour == late_hour
		if start_time.min > late_minute
			puts "#{start_time}\tLate"
		end
	end

	working_hour = ((end_time.hour * 60 + end_time.min) - (start_time.hour * 60 + start_time.min)) / 60.0
	if working_hour < 9.0
		puts "#{end_time}\tOnly #{working_hour.round(2)} hours"
	end
end

def normal_report(xls)
    target_month = 5
    target_year = 2011
    last_date_time = nil
	start_time = nil
	end_time = nil
	
	xls.last_row.downto(2) do |line|
		date_time_s = xls.cell(line,'D').to_s.split
		# date
		date_s = date_time_s[0]
        date = Date.parse(date_s)
        
		# if year and month are not match, continue loop
        next if date.month != target_month || date.year != target_year
        
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
				calculate_working_hours(start_time, end_time)
			end
			# store start_time
			end_time = nil
			start_time = date_time
		end
		
		# calculate the calculate_working_hours on the last day like 5/31
		if line == 2
			end_time = date_time
			calculate_working_hours(start_time, end_time)
		end
		
		last_date_time = date_time
	end
end

def intern_report(xls)
	puts 'This is intern!'
end

def generate_report(xls, intern)
	name = xls.cell(2,'B')
	puts "#{name}:"
	
	if intern.include?(xls.cell(2,'A').to_i)
		intern_report(xls)
	else
		normal_report(xls)
	end
end

#define intern employee number
intern = [23]

Dir.glob("#{pwd}/*.xls") do |file|
  file_path = "#{pwd}/#{file}"  
  file_basename = File.basename(file, ".xls")  
  xls = Excel.new(file_path)
  #xls.to_csv("#{pwd}/#{file_basename}.csv")
  generate_report(xls, intern)
end
