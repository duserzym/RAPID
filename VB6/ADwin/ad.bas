'<ADbasic Header, Headerversion 001.001>
' Process_Number                 = 1
' Initial_Processdelay           = 1000
' Eventsource                    = Timer
' Control_long_Delays_for_Stop   = No
' Priority                       = High
' Version                        = 1
' ADbasic_Version                = 5.0.6
' Optimize                       = Yes
' Optimize_Level                 = 1
' Info_Last_Save                 = PYRITE  PYRITE\lab
'<Header End>
#include adwingoldii.inc
#include globals.inc

#define pd 12000 ' 100µs = 10 kHz
#define ec (10/3)*1E-9 ' execution cycle 3.3ns

#define tc ec*pd ' time constant = process delay in seconds

#define noiselevel 100

DIM cnt as LONG
DIM bufn as LONG
DIM sign as LONG
DIM acval as LONG

lowinit:
  processdelay = pd '300 = 1e-6s
  bufn = 1 ' number of buffer to be filled
  BUF1_FULL = FALSE ' buffer 1 empty
  BUF2_FULL = FALSE ' buffer 2 empty
  MAXACVAL = 0
  MINACVAL = 0

  cnt = 1

  PAR_80 = 0
  PAR_81 = 0


  DCPHASE = DCPHASE_READY

event:
  acval = ADC( PORT_ACCUR) - 32768 ' read ADC bipolar
  DCVAL = ADC( PORT_DCCUR) - 32768 ' read ADC bipolar	

  ' record dc values
  'IF( dccnt < 100000) THEN
  '	DCVALS[ dccnt] = dcval
  '	dccnt = dccnt + 1
  'ENDIF
	
	
  IF( (acval < noiselevel) AND (acval > -noiselevel)) THEN ' reset maximum and minimum acvalues when signal close to zero
    IF( AMPSIDX < NAMPS) THEN ' record amplitudes until arrays are full
      IF( MAXACVAL > noiselevel) THEN
        IF( (MAXACVAL > PEAKVOLTAGE * MAX16BIT/20) AND (ACPHASE = ACPHASE_RAMP_UP)) THEN 
          ' when we ramp up and reach the peak field
          ACPHASE = ACPHASE_PEAK_REACHED
        ENDIF
				
        IF( ((MAXACVAL < DC_OFF * MAX16BIT/20) AND (ACPHASE = ACPHASE_RAMP_DOWN)) AND (DCPHASE = DCPHASE_ON)) THEN 
          ' if we ramp down and reach the specified field, switch off DC field
          DAC( PORT_DCCUR, MAX15BIT) ' set output to zero
          DCPHASE = DCPHASE_DONE
        ENDIF
	
        IF( ((MAXACVAL < DC_ON * MAX16BIT/20) AND (ACPHASE = ACPHASE_RAMP_DOWN)) AND (DCPHASE = DCPHASE_READY)) THEN 
          ' if we ramp down and reach the specified field, switch on DC field
          DAC( PORT_DCCUR, MAX16BIT/20 * DCFIELD + MAX15BIT) ' output specified voltage 
          DCPHASE = DCPHASE_ON ' dc field is switched on
        ENDIF
      ENDIF
    ENDIF
  ENDIF

  ' update maximum and minimum value in this oscillation
  MAXACVAL = MAX_LONG( MAXACVAL, acval)
  MINACVAL = MIN_LONG( MINACVAL, acval)


  if( (bufn = 1) AND (BUF1_FULL = FALSE)) THEN ' fill buffer1
    BUF1[ cnt] = acval
    cnt = cnt + 1
    PAR_80 = PAR_80+1
  ENDIF

  if( (bufn = 2) AND (BUF2_FULL = FALSE)) THEN ' fill buffer2
    BUF2[ cnt] = acval
    cnt = cnt + 1
    PAR_81 = PAR_81+1
  ENDIF

  if( cnt > POINTS) THEN ' buffer full, switch to next
    cnt = 1
    if( bufn = 1) THEN
      BUF1_FULL = TRUE ' buf1 full, will be reset by other process
      bufn = 2
    ELSE
      IF( bufn = 2) THEN
        BUF2_FULL = TRUE
        bufn = 1
      ENDIF
    ENDIF
  ENDIF


