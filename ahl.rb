require 'win32ole'

$work = 0
$ticket = "" #ticket number
$tickets = 0
$t = 0
$tt = 0 #total time

def str_time(time)
	"#{Time.now.strftime("%H:%M")} #{time-2} ticks : #{time > 60 ? time/60 : '0'} minutes have passed"
end

def refresh
	s =WIN32OLE.new("WScript.Shell")
	$t = 0
	p "Focus on window you intend to refresh."
	reaper = Thread.new do 
		while true
			r = rand(60)
			$t += r
			p "Refreshing in: "
			r.times do 
				r -= 1
				p r
				sleep 1
			end
			p "Refreshed at #{str_time($t)} since refreshing started."
			s.SendKeys("^r")
		end
	end
	input = gets.chomp
	reaper.kill
	exit input
		
end

def ticket
	$t = 0
	p "Work began at #{Time.now.strftime("%H:%M")}."
	work = Thread.new do
		while true
			$t += 1
			p $t
			sleep 1
		end
	end
	input = gets.chomp
	p "Work ended at #{str_time($t + 2)} since work started."
	work.kill
	$tt += $t
	$tickets += 1
	exit input
end

def exit command
	case command
	when "s"
		p "#{$tt} seconds (#{$tt/60} minutes) in total work time. #{$tickets} tickets completed."
		exit "p"
	when "p"
		p "Paused. Enter a new command to continue."
		input = gets.chomp
		exit input
	when "q"
		p "Exiting Program"
	when ""
		#ticket

		if $work == 0
			$work = 1
			p "Ticket #:"
			$ticket = gets.chomp
			p "Working on Ticket #{$ticket}"
			ticket
			
		else
			$work = 0
			p "Completing Ticket #{$ticket}"
			sleep 1
			exit "p"
			
		end
	when "r"
		#refresh
		refresh
	when "test"
		begin
		system("lpr", "text.txt") or raise "lpr failed"
		rescue StandardError => e
			puts e
		end
	else
		p "You said #{command}, but for some reason that wasn't recognized."
		p "Please enter another command."
		input = gets.chomp
		exit input
	end

end

p "Welcome to Automation Heavy Lifter"
exit "p"

