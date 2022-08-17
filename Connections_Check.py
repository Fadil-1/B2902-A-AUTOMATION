"""
Written by Fadil Isamotu (v1.0)
November 17, 2021
moisa1@morgan.edu
"""
import pyvisa
from time import sleep as idle 

#------------------------------------------------------------------------------------------------PYVISA INITIALIZATION-------------------------------------------------------------------------------------------------
# pyvisa's resource manager to get devices Id
rm = pyvisa.ResourceManager()
#print(rm.list_resources()) #__Optional__ To check what visa devices are available
mtrx = rm.open_resource('USB0::0x0957::0x3D18::MY56480010::0::INSTR') # Switching matrix ID
smu  = rm.open_resource('USB0::0x0957::0x8C18::MY51140489::0::INSTR') # SMU ID

#-----------------------------------------------------------------------------------------------------SWITCHING MATRIX----------------------------------------------------------------------------------------------------
def matrix_Open(device):
    """This function opens specified channels on DB25 Breakout 25 Pin Connector using U2751A switching matrix.  
    
    >>> matrix_Open(1)
    Opens channels for transistor 1

    >>> matrix_Open("all")
    Opens all channels

    >>> matrix_Open("RST")
    Resets the switching matrix (Opens all channels)
    """  
    
    if device == 'all' or device == 'ALL' or device == 'All':
        
        mtrx.write(f'ROUTe:OPEN (@ 101, 202, 103, 204, 105, 206, 107, 208)')
        
    elif device == "RST" or device == "RESET" or device == "reset" or device == "rst":
        
        mtrx.write('*RST')    
        
    elif device == '1':
        
        mtrx.write(f'ROUTe:OPEN (@ 101, 202)')
        
    elif device == '2':
        
        mtrx.write(f'ROUTe:OPEN (@ 103, 204)')
        
    elif device == '3':
        
        mtrx.write(f'ROUTe:OPEN (@ 105, 206)')
    
    elif device == '4':
        mtrx.write(f'ROUTe:OPEN (@ 107,208)')

    mtrx.write('opc?')                       # Waits for all commands to be executed before executing next commands   

def matrix_Close(device): 
    """This function closes specified channels on DB25 Breakout 25 Pin Connector using U2751A switching matrix.  
    
    >>> matrix_Close(1)
    Closes channels for transistor 1

    >>> matrix_Close("all")
    Closes all channels

    >>> matrix_Close("RST")
    Resets the switching matrix (Opens all channels)
    """ 
    
    if device == 'all' or device == 'ALL' or device == 'All':
        
        mtrx.write(f'ROUTe:CLOSE (@ 101, 202, 103, 204, 105, 206, 107, 208)')
        
    elif device == "RST" or device == "RESET" or device == "reset" or device == "rst":
        
        mtrx.write('*RST')    
        
    elif device == '1':
        
        mtrx.write(f'ROUTe:CLOSE (@ 101, 202)')
        
    elif device == '2':
        
        mtrx.write(f'ROUTe:CLOSE (@ 103, 204)')
        
    elif device == '3':
        
        mtrx.write(f'ROUTe:CLOSE (@ 105, 206)')
    
    elif device == '4':
        mtrx.write(f'ROUTe:CLOSE (@ 107,208)')
    
    mtrx.write('opc?')                       # Waits for all commands to be executed before executing next commands    


def constant_Voltage(c1 = 5, c2 = 5, I1_Compliance = 100e-3, I2_Compliance = 100e-3, idle_Time = 15, channel = "All"):
    """This function sets a constant voltage as well as a limit current from both channels of A B2902A SMU.

    >>> constant_Voltage(10, 5, 1, 0.2)
    Channel one outputs 10 volts with a 1A current compliance
    Channel two outputs 5  volts with a 2A current compliance
    """

    smu.write('*rst')                                   # Resets the sourcemeter
    mtrx.write('*RST')                                  # Resets the switching matrix

    matrix_Close(channel)
                                                  
    smu.write(":SOUR1:FUNC:MODE VOLT; :SOUR2:FUNC:MODE VOLT")                        # Sets channel 1 and 2 to source voltage
    smu.write(f':sens1:func ""curr""; :sens2:func ""curr""')                         # Sets channel 1 and 2 to sense current
    smu.write(f':sens1:curr:prot {I1_Compliance}; :sens2:curr:prot {I2_Compliance}') # Setscurrent compliance for channel 1 and 2
    smu.write(f':SOUR1:VOLT {str(c1)}; :SOUR2:VOLT {str(c2)}')                       # Sets voltage values for channel 1 and 2
    smu.write(":OUTP1 ON; :OUTP2 ON")                                                # Turns on channel 1 and 2  
    smu.write('opc?')                                                                # Waits for all commands to be executed before executing next commands
    
    idle(idle_Time)
    smu.write('*rst')                                   # Resets the sourcemeter
    mtrx.write('*RST')                                  # Resets the switching matrix

constant_Voltage()