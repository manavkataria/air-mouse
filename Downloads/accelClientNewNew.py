#!/usr/bin/python

import socket
from math import sqrt, fabs
from Xlib import X, display


d = display.Display()
s = d.screen()
root = s.root
root.warp_pointer(1280,770)
d.sync()


lsock = socket.socket()
lsock.connect(("192.168.0.202",7000))
lfd = lsock.makefile('r')
print "connected"

xo = 41;
yo = 41;

deadZone_x = 100;
deadZone_y = 90;

x_speed_scalor = 0.01;
y_speed_scalor = 0.008;

while 1 != 0 :
#    acceleration = lsock.recv(100)[1:-1].split(',')
#    print (acceleration[0])+' '+(acceleration[1])+' '+(acceleration[2])
    pointer = root.query_pointer()
    xCur = pointer.root_x
    yCur = pointer.root_y

    in_str = lfd.readline()
    print in_str
    v = in_str.split(',')
    x = int(v[0])
    y = int(v[1])
    z = int(v[2])
    try:
        xsign = (x)/fabs(x)
    except:
        xsign = 0
    try:
        ysign = (y)/fabs(y)
    except:
        ysign = 0
 
#    sum = int( sqrt(x*x+y*y+z*z))
#    print x, y,

    if fabs(x) > deadZone_x:
#        print "__X__",x,
        #xCur = pointer.root_x+2*xsign;
        xCur = pointer.root_x+x_speed_scalor*x;
        
   
    if fabs(y) > deadZone_y:
#        print "__Y__",y,
        #yCur = pointer.root_y-1*ysign;
        yCur = pointer.root_y+y_speed_scalor*y;
   
#    print x,y,xCur,yCur,xsign,ysign
    
    root.warp_pointer(xCur, yCur)
    d.sync()



