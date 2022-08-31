# B2902_A + SWITCHING MATRIX AUTOMATION

This program was initially implemented to do custom measurements for a research project. It is still very "Case Specific" as it works with a select type of instruments. Nontheless it'd be reltively simple to "retune" for different end-goals.
The functions from the main script can be copied and used independently with the instruments they are meant to control or, in the case of the SMU related functions, with some slighly older or newer models of Agilent/Keysight SMUs.

The setup contains three fundamental instruments:

A sourcemeter (Agilent B2902-A);
A breakout board (db25 breakout board);
A switching matrix (Keysight U2751A);

The research this setup was made for required the DUTs to be far from the instruments, therefore all three instruments were connected through USB using a data cable and a USB extender(See image below). 


![2](https://user-images.githubusercontent.com/89228814/185051264-d0e14fa6-b6f3-4135-92b1-5567876989df.JPG)

For pineout discription the DUTs are connected on a breadboard and powered with color-coded wires. 


![3](https://user-images.githubusercontent.com/89228814/185051668-db693824-718c-48b6-b995-fa2f48109c7d.JPG)

The full setup is shown in the image below:
![set](https://user-images.githubusercontent.com/89228814/185050556-d3a23d2e-c06f-4f88-b013-ec63a2118b85.JPG)

The front “high” output of the sourcemeter serves as drain bias, while the rear “high” output serves as the gate bias. Red represents gate bias, yellow represents drain bias, and black represents source/ground.

Match all wires from the breakout board to the breadboard from left to right. The rightmost set of wires connects a diode, the middle connects transistor_2/DEVICE_2, and the leftmost connects transistor_1/DEVICE_1. The yellow wire can be omitted (gate wire) for the diode, as it is a two-terminal device.

![IMG_5091 (2)](https://user-images.githubusercontent.com/89228814/185053369-c442a8e1-020c-4233-91ee-6fab4ff3ea04.JPG)


This setup is designed to work with 3 devices, namely, a diode, and two transistors.

The program was hardcoded(Does not use the B2902A SMU's built-in sweep functions) using python and Standard Commands for Programmable Instruments (SCPI) to perform IV, ID/VD and ID/VG sweeps using the channels of a B2902A SMU. The sweeps output are written to a CSV file within a directory dynamically specified by the "directory" parameter in a input JSON file. Other Json options are provided to tune the voltages, current limit, measurement's granularity and much more.


