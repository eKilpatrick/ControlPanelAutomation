# Control Panel Automation

The goal of this project is to use robotic arms to automate the building of control panels for circuit breakers. There are many challenges involved in this project and it has been ongoing for multiple years with many different groups of people working on it.

Oracle DB Setup for this Project
--------------------------------
Tables for:
	Part Sizes
	Storage of Parts
	Location of Parts
	Past Orders

Every circuit breaker has an Inventor 3D file associated with it that details where on the control panel all of the parts are supposed to go.
When a breaker's order number is entered into this program, that Inventor file is pulled up and the coordinate data for the parts is extracted and then saved into the 
database. This Inventor extraction was written by a 3rd party AutoDesk programming company. Next, the program loops through all of the parts in that order. When it gets 
to a part that hasn't been placed/screwed onto the panel yet, it starts working on that part. It gets the location of that part along with the size of that part. All of 
the necessary variables associated with that part are then calculated. Next, these variables are saved to the robot's input registers via the Real Time Data Exchange 
port. The program on the robots are started. The gripper robot goes to the storage bin and grabs the part, meanwhile the screwdriving robot goes to the screw presentor 
and gets a screw. Once the gripper has the part, it goes to the panel and places it in the correct spot. Once the part is placed, the screwdriver screws in the first 
screw and goes to get the second. The gripper goes back to its home position and the screwdriver finishes screwing in the part. This process is repeated for every part 
until that order is completed. An operator will then roll their cart out of the area and the next operator will roll theirs in, and go through this same process.
