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

'#define POINTS 256 
'#define POINTS3 768
'#define POINTS4 4096
'#define POINTS05 128
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
#define ACPHASE_PEAK_REACHED 2
#define ACPHASE_PEAK_DELAY 3
#define ACPHASE_RAMP_DOWN 4
#define ACPHASE_DONE 5

'--------------------------------------------------------------------------------
'Modifications for different styles of Ramp cycle needed by the Visual Basic
'GUI to do:
' 1 - Normal, Monitored AF RAMP, without external storage of Ramp data
' 2 - Debug mode, Monitored AF RAMP, with external storage of Ramp data
'     this mode will also be used to do Field Calibration of the AF Ramp in tandem
'     with a computer controlled Hirst 908A Guassmeter & Axial or Transverse Hall-probes
' 3 - Clipping test, unmonitored AF RAMP, with external storage of Ramp data
'--------------------------------------------------------------------------------
#define AFRAMP_MONITORED 1
#define AFRAMP_DEBUG 2
#define CLIPTEST 3
'#define AFTUNE 4

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
#define PEAK_DELAY  FPAR_34 ' seconds
#define FREQ FPAR_35 ' frequency of ac field in Hz
'#define FREQMAX FPAR_36
#define AC_AMPL_LIMIT  FPAR_36 ' volts (hard limit of AF amplitude)


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
'
'Because these numbers are used to size arrays, they have to be constants
'SineFit process runs every 3 ms, so 10000 sinefits = 30 secs
'Monitor process runs every 10 micro-s, so 3000000 monitor pts = 33.33 secs
'Ramp Output process runs every 3.3 micro-s, so 10000000 = 30 secs
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
#define NOISELEVEL PAR_35
'--------------------------------------------------------------------------------
'Digital port settings for controlling TTL relays
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'Modifications for declaring external memory global data arrays for saving the AF Ramp data
'and the sine fit parameters for them to be retreived by the external program GUI
'--------------------------------------------------------------------------------
'External Parameters for transfering Ramp Data to external GUI code
#define AFRAMPDATA DATA_31      'AF Ramp AC monitor input, Ramp input point #
#define AFRAMPOUTPUT DATA_32    'AF Ramp AC Output, Ramp output point #
'#define SINEFITRESULTS DATA_33  'Sine Fit results - 4 parameters, RMS, start point


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
'#define SINEFITCOUNT PAR_11

'Flag status variables to tell when LOCAL ramp data arrays are full and should
'be read into the external variables
'#define RAMPDATA_1_FULL PAR_12
'#define RAMPDATA_2_FULL PAR_13
'#define ARMINPUT_1_FULL PAR_15
'#define ARMINPUT_2_FULL PAR_16
'#define ARMOUTPUT_1_FULL PAR_17
'#define ARMOUTPUT_2_FULL PAR_18

'--------------------------------------------------------------------------------
'Modifications for defining global vars for use as status flags for the sine-fit
'process
'--------------------------------------------------------------------------------
'#define SINEDATA_FULL PAR_19
'#define SINEFIT_DONE PAR_20
'#define DOSINEFIT PAR_21

'This variable is to let all the processes know if there was a fatal
'inversion error during the last sine-fit and to use the last sine-fit
'values until the next sine-fit is run
'#define INVERSIONERROR PAR_22
'
''This variable is to store the start point (relative to the analog input LC circuit
''monitor process) of the first point in the sine-fit local data array.  This value
''will be saved to the external memory sine-fit results data array.
'#define SINEFITSTARTPT PAR_23
'
''Flag to indicate that the RampUp slope has been reduced by the analog input
''signal monitor process in monitor.bas
'#define RAMPSLOWED PAR_24

'--------------------------------------------------------------------------------
'Modifications for declaring global data arrays necessary for doing the sine-fit
'and making the sine-fit results available to other board processes
'--------------------------------------------------------------------------------
'Internal ADBasic code parameters for running sinefit
'#define SINEDATA DATA_1
'#define SINEFITPARAMS DATA_2

'--------------------------------------------------------------------------------
'Local array variables to store up to 1028 points of Ramp input and output
'that will be transfered to AFRAMPDATA & AFRAMPOUTPUT after every 1028 points
'are written to the local variables
'--------------------------------------------------------------------------------
'#define LOCALRAMPDATA_1 DATA_3
'#define LOCALRAMPDATA_2 DATA_4
'#define LOCALRAMPOUTPUT_1 DATA_5
'#define LOCALRAMPOUTPUT_2 DATA_6
'#define LOCALARMINPUT_1 DATA_7
'#define LOCALARMINPUT_2 DATA_8
'#define LOCALARMOUTPUT_1 DATA_9
'#define LOCALARMOUTPUT_2 DATA_10

'********************************************************************************
'********************************************************************************
'
'
'**********************  Global Variable Declarations  **************************
'
'
'********************************************************************************

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'Modifications
'(March 2010, I Hilburn)
'
'Modifications are to create global variables needed for the Sine-Fit process
'and to enable monitoring the AF Ramp LC-circuit sine-wave amplitudes through
'fitting 256 points worth of analog input signal data 
'--------------------------------------------------------------------------------

''Parameters determined by last sine fit
'Dim SINEFITPARAMS[5] as FLOAT at DM_LOCAL
'
''Arrays needed to transfer data to the Sine-Fit algorithm
'Dim SINEDATA[POINTS] as LONG at DM_LOCAL
'
'Dim SINEFITRESULTS[10000][6] as FLOAT at DRAM_EXTERN

'External array (GIANT!!) for storing raw analog input signal for AF Ramp
Dim AFRAMPDATA[NumMonitorPts] as LONG at DRAM_EXTERN

'External array (GIANT!!) for storing the raw analog output signal for AF Ramp
Dim AFRAMPOUTPUT[NumMonitorPts] as LONG at DRAM_EXTERN
  
'Internal Versions of these arrays that will be incrementally saved to external arrays above
'Dim LOCALRAMPDATA_1[POINTS4] as LONG at DM_LOCAL
'Dim LOCALRAMPDATA_2[POINTS4] as LONG at DM_LOCAL
'Dim LOCALRAMPOUTPUT_1[POINTS4] as LONG at DM_LOCAL
'Dim LOCALRAMPOUTPUT_2[POINTS4] as LONG at DM_LOCAL
'Dim LOCALARMINPUT_1[POINTS4][2] as LONG at DM_LOCAL
'Dim LOCALARMINPUT_2[POINTS4][2] as LONG at DM_LOCAL
'Dim LOCALARMOUTPUT_1[POINTS4][2] as LONG at DM_LOCAL
'Dim LOCALARMOUTPUT_2[POINTS4][2] as LONG at DM_LOCAL



'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
