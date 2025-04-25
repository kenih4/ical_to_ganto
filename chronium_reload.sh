#!/bin/sh

#	ps aux | grep getscreenshot_loop.sh | grep -v grep | awk '{ print "kill -9", $2 }' | sh | sleep 3 | sh

	host=`hostname`
	echo host = ${host}

	while :
	do
		now_date=`date  +"%Y%m%d-%H%M"`	#"%Y%m%d-%H%M%S"`
		echo "now_date = 	[${now_date}]"

#		firefox http://saclaopr19.spring8.or.jp/~lognote/calendar/gantt-group-tasks-together.html

#		chromium-browser http://saclaopr19.spring8.or.jp/~lognote/calendar/gantt-group-tasks-together.html

		xdotool search --onlyvisible --class "chromium" windowactivate key F5

#		sleep 1 ; xwd -root | convert - /home/xfel/xfelopr/kenichi/screenshot_loop/$now_date-$host.png
#		sleep 1 ; cp /home/xfel/xfelopr/kenichi/screenshot_loop/$now_date-$host.png /home/xfel/xfelopr/kenichi/screenshot_loop/$host.png
#		find /home/xfel/xfelopr/kenichi/screenshot_loop -type f -name "*.png" -mtime +90 | xargs rm -f
		
		sleep 3600	#3600	#sec
	done
