REBOL[MY TEST]
System/ports/serial: [ com3 ] ;if using linux this line is slightly different
ser: open/direct/no-wait serial://port1/250000/none/8/1 
ser/rts-cts: false
update ser
buffer: make string! 1
logFile: "Hx19log.txt"
xdata: array/initial 6 0
bb: 0 bc: 0 tt: "" kt: "" nm: 1 kk: "" bd: 0
zz: -30
;***************************************** CONDITIONS ********************************
;This program utilizes equations found at www.hexamite.com/hx19posb.htm. For this 
;set of relationship to work, line distance from first receiver to the second presents 
;the X axis, and line distance from first receiver to third represents the Y-axis. And
;it follows that there is a 90 degree angle between X, Y and Z axis. 
;-------------------------------------------------------------------------------------
;Following function computes xx string checksum and appends /hexsum to the string
;where hexsum is the hexadecimal representation of the sum of the string characters
hx19: funct [xx] 
   [
     pp: copy xx
     ss: 0
     foreach char pp [ss: ss + char]
     hx: to-hex ss
     until [sn: take hx sn > #"0"]
     insert hx sn
     append pp "/" append pp hx append pp CR
     return pp
   ]
;following function sets the display layout with check buttons and etc.
view/new/title layout 
  [
    across 
	label "SYNC"
	check [
               either bb > 0 
		[bb: 0 insert ser hx19 "M&%" update ser]
		[bb: 1 insert ser hx19 "M&$" update ser]
               ]
	label "LOG"
	check [
                either bc > 0
                 [bc: 0 close log]
                 [bc: 1 log: open/new %hx19log.txt]
               ]
	label "STOP"
 	check [
                either bd > 0 
                 [bd: 0]
                 [bd: 1]
               ]

     f: field 200
   	btn "TX" [insert ser hx19 f/text update ser]
	btn "SetUP File" 
	 [
	   px: read/lines %hx19setup.txt
           repeat nn 20
            [
             kk: pick px nn
             if 0 = length? kk [break]
             print kk
             insert ser hx19 kk update ser
             wait 0.05
            ]
         ]
  ] "HX19 ACCESS"
insert-event-func [switch event/type [close [close ser if bc > 0 [close log] quit]]event]
;the following code monitors incoming data from a serial port and prints in a popup box
forever 
[
until 	 [
	  while[empty? buffer][read-io ser buffer 1 wait 0.0001]
	  x: to-integer first buffer
	  append tt buffer
	  clear buffer
	  x = 13
	 ]

   either (length? tt) > 13 
    [
     kx: copy/part (at tt 10) length? tt
     foreach char kx 
      [
       if char < 33 [break]
       kk: append kk char
      ]
     poke xdata nm to decimal! kk
     kt: copy/part (at tt 5) 3
     nm: nm + 1
     clear kk
    ]
    [
;print ["tag" kt " R1:" xdata/1 " R2:" xdata/2 " R3:" xdata/3] 
;round(square-root zz)]]
      x0: (xdata/1) * (xdata/1)
      x1: (xdata/2) * (xdata/2)
      x2: (xdata/3) * (xdata/3)		;tag to receiver distance array
      xx: ((x0 - x1) + 1000000) / 2000
      yy: ((x0 - x2) + 1000000) / 2000  ;find these equations explained on the hexamite website
      zz: (x0 - (xx * xx) - (yy * yy))
      if bd = 0 
	[
     	 if zz > 2 [print ["tag" kt " X:" round(xx) " Y:" round(yy) " Z:" round(square-root zz)]]
	]
      nm: 1
    ]
    if bc > 0 [append log tt]
    clear tt
]
