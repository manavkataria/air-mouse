/*
 * Accelerometer Mouse
 * by Manav Kataria <http://air-mouse.googlecode.com>
 * 6th May 2009
 * 
 * Sends accelerometer data to the computer via serial port. 
 * Uses a windows driver WinDriver to move the mouse on windows. 
 * Requires MSCOMM32.ocx (ActiveX control)
 
 * Updated: July 2009 
 * Trying to reduce delay
 */

/* Enables Z Axis information to be txd to PC Packet */
//#define DEF_3AXIS

#include "WProgram.h"
void setup();
void loop();
int xpin = 0;    // select the input pin for x-axis of the accelerometer
int ypin = 2;
int zpin = 1;
int leftpin = 19;
int rightpin = 18;

void setup() {
  Serial.begin(115200);
  pinMode(leftpin, INPUT);
  pinMode(rightpin, INPUT);  

}

void loop() {
   //Serial.write(analogRead(zpin>>2));     delay(1);
   //Serial.print(" ");
  
  int marker=0xAA;
  int left=0, right=0;

  left  = (digitalRead(leftpin) <<1);      delay(5);
  right = (digitalRead(rightpin)<<0);      delay(5);
  
  Serial.write(marker);                    delay(5);
  //Serial.print("x: ");                    delay(5);
  Serial.write(analogRead(xpin)>>3);       delay(5);
  //Serial.print(" y: ");                    delay(5);
  Serial.write(analogRead(ypin)>>3);       delay(5); 
#ifdef DEF_3AXIS
  //Serial.print(" z: ");                    delay(5);
  Serial.write(analogRead(zpin)>>3);       delay(5);
#endif
  Serial.write(left|right);                delay(5); 

}

int main(void)
{
	init();

	setup();
    
	for (;;)
		loop();
        
	return 0;
}

