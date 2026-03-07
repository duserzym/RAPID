'<ADbasic Header, Headerversion 001.001>
'<Header End>
'********************************************************************************
'********************************************************************************
'
'
'****************************  Global Constants  ********************************
'
'
'********************************************************************************

#define PI 3.14159265358979

'For 20 Volt range, and 16-bit data conversion, this is the least significant bit
'the size of the smallest increment that the voltage can be read in / written out at
#define LSB 0.00030517578125  'Volts
#define MAX16BIT 65536        'Largest ADWIN count value = +10 Volts
#define MAX15BIT 32768        'Median ADWIN count value = 0 Volts

#define TRUE 1
#define FALSE 0

'RAMP Phase Constants
#define ACPHASE_START 0
#define ACPHASE_RAMP_UP 1
#define ACPHASE_OVERRAMP 2
#define ACPHASE_PEAK_REACHED 3
#define ACPHASE_PEAK_DELAY 4
#define ACPHASE_RAMP_DOWN 5
#define ACPHASE_RAMP_ENDING 6
#define ACPHASE_DONE 7

'Ramp Down method
#define RAMPDOWN_SLOPE 0
#define RAMPDOWN_PERIOD 1

'--------------------------------------------------------------------------------
'Modifications for different styles of Ramp cycle needed by the Visual Basic
'GUI to do:
' 1 - Normal, Monitored AF RAMP, without external storage of Ramp data
' 2 - Debug mode, Monitored AF RAMP, with external storage of Ramp data
'     this mode will also be used to do Field Calibration of the AF Ramp in tandem
'     with a computer controlled Hirst 908A Guassmeter & Axial or Transverse Hall-probes
' 3 - Clipping test, unmonitored AF RAMP, with external storage of Ramp data
' 4 - AF Tuning test, unmonitored AF Ramp, held at peak, with varying freq and external
'     storage of Ramp data
'--------------------------------------------------------------------------------
#define AFRAMP_MONITORED 1
#define AFRAMP_DEBUG 2
#define CLIPTEST 3
#define AFTUNE 4

'********************************************************************************
'********************************************************************************
'
'
'*************************  externally set parameters  **************************
'
'
'********************************************************************************

#define SLOPE_UP  FPAR_31  ' volts / second
#define SLOPE_DOWN  FPAR_32  ' volts / second
#define PEAKVOLTAGE  FPAR_33 ' volts
#define FREQ FPAR_34 ' frequency of ac field in Hz
'#define FREQMAX FPAR_36
#define AC_AMPL_LIMIT  FPAR_35 ' volts (attempted limit of AF amplitude)
#define MAX_RAMPVOLTAGE FPAR_36 ' volts - absolute max that ramp voltage can go up to
'         before the ramp is terminated
#define MAX_PEAKVOLTAGE FPAR_37 ' volts - absolute max that the input peak voltage can go up
'         to before the ramp is terminated


'--------------------------------------------------------------------------------
'Modifications for different styles of Ramp cycle needed by the Visual Basic
'GUI to do:
' 1 - Normal, Monitored AF RAMP, without external storage of Ramp data
' 2 - Debug mode, Monitored AF with external storage of Ramp data
'     this mode will also be used to do Field Calibration of the AF Ramp in tandem
'     with a computer controlled Hirst 908A Guassmeter & Axial or Transverse Hall-probes
' 3 - Clipping test, unmonitored AF RAMP, with external storage of Ramp data
' 4 - AF Tuning test, unmonitored AF Ramp, held at peak, with varying freq and external
'     storage of Ramp data
'--------------------------------------------------------------------------------
#define RAMPMODE PAR_31

'--------------------------------------------------------------------------------
'Modifications for pulling in the number of points that external data arrays should have
'allocated to them for sending AF ramp data and sine fit data to external program GUI
'--------------------------------------------------------------------------------
#define NumMonitorPts 1000000

'--------------------------------------------------------------------------------
'Modified values of three ports to be setable by passing in a parameter
'from the external GUI program.
'--------------------------------------------------------------------------------
#define PORT_SINEOUT PAR_32  '(March 2010, I Hilburn - creating global const to store the DAC-OUT port number for the sine wave output)
#define PORT_ACCUR PAR_33

'--------------------------------------------------------------------------------
'Speed in Hz at which to run the AF Monitor, and Sine Output processes
'--------------------------------------------------------------------------------
#define AFRAMP_PD PAR_34

'--------------------------------------------------------------------------------
'Noiselevel for 16-bit count comparisons
'In 16-bit long data type counts (0 - 65536)
'--------------------------------------------------------------------------------
#define NOISELEVEL PAR_35

'--------------------------------------------------------------------------------
'Peak Delay time in Periods / Cycles
'--------------------------------------------------------------------------------
#define PEAKDELAY_PERIODS PAR_36

'--------------------------------------------------------------------------------
'Variables to Set the number of periods / cycles for the Ramp Down
'and to select between using the Ramp Down slope or using the number of periods
'
'RAMPDOWN_MODE = 0 -- use slope (Volts / second) to ramp down
'RAMPDONW_MODE = 1 -- use number of periods to ramp down
'--------------------------------------------------------------------------------
#define NUMPERIODS PAR_37
#define RAMPDOWN_MODE PAR_38


'--------------------------------------------------------------------------------
'Modifications for declaring external memory global data arrays for saving the AF Ramp data
'and the sine fit parameters for them to be retreived by the external program GUI
'--------------------------------------------------------------------------------
'External Parameters for transfering Ramp Data to external GUI code
#define AFRAMPDATA DATA_31      'AF Ramp AC monitor input, Ramp input point #
#define AFRAMPOUTPUT DATA_32    'AF Ramp AC Output, Ramp output point #

'********************************************************************************
'********************************************************************************
'
'
'*************************  internally set parameters  **************************
'
'
'********************************************************************************

'Amplitude of AC sine signal output in volts
#define SINEAMP FPAR_2 

'Monitor Peak variables
#define ACVOLT FPar_3
#define RAMPMAX FPar_4
#define OUTMAX FPar_5

'Time Steps for various processes
#define ACOUT_TIMESTEP FPAR_6

'Period of one oscillation in output points - float
#define NSAMPLES FPar_7

'AC Input and Output for monitored AF Ramp process
#define ACOUTPUT PAR_2
#define ACINPUT PAR_3

#define ACPHASE  PAR_4  ' phase of process (ramp up, peak, down)

'--------------------------------------------------------------------------------
'Modifications in scope for internal use parameters (all parameters below PAR_31)
'These are READ-only values for the external program.
'--------------------------------------------------------------------------------
'Global counter variables for counting how many ADC or DAC points have been input
'and output in the AF-Ramp process.
#define OUTCOUNT PAR_5
#define INCOUNT PAR_6
#define RAMPUPLASTPT Par_7
#define RAMPDOWNFIRSTPT Par_8
#define MAXALLOWEDDATAPTS Par_11

'********************************************************************************
'********************************************************************************
'
'
'**********************  Global Variable Declarations  **************************
'
'
'********************************************************************************

'External array (GIANT!!) for storing raw analog input signal for AF Ramp
Dim AFRAMPDATA[NumMonitorPts] as LONG at DRAM_EXTERN

'External array (GIANT!!) for storing the raw analog output signal for AF Ramp
Dim AFRAMPOUTPUT[NumMonitorPts] as LONG at DRAM_EXTERN

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
