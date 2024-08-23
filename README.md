# RPGWorkOrderCutter
This was an internal project at Railway Products Group. The program records coordinates for various text boxes and buttons on a Sage100 menu. It has to be initially trained (points saved in a text file) before it can be ran automatically. 

Once trained, the program will read the part number, quantity and job number from the text boxes and proceed to loop through the saved coordinates until there are no more part numbers/quantities/job numbers to use. 

The goal of this small project was to significantly speed up repetitive tasks someone within the company had to do on a daily basis. 

The main libraries used for this project were PyAutoGUI for the mouse coordinate logic (moving, clicking, saving) and Tkinter for the user interface implementation. The mouse library is also used to wait for mouse clicks so that coordinates could be saved. 
