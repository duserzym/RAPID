Paleomag 2010 2.3.9

by Laurent Carporzen & Bob Kopp & Scott Bogue & Isaac Hilburn
January 27th, 2010
Licensed under the terms of the GNU General Public License

*****

Contents:
	1. Requirements
	2. Installation
	3. Component Files

*****

1. REQUIREMENTS

matRockmag was written in Visual Basic 6 and has been tested under Windows 98,
Windows 2000, and Windows XP. 

For email, it uses vbSendMail by Dean Dusenbery and FreeVBCode.com, available at
http://www.freevbcode.com/ShowCode.Asp?ID=109 and included within the accompanying vbSendMail.zip.

By default, it attempts to read and write its INI file from C:\Paleomag\Paleomag.INI.
This location can be changed by editing
MyComputer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Paleomag\INIFile. If
you do not wish to do this, you should ensure the directory C:\Paleomag exists.

Software for configuring Measurement Computing cards (used to communicate with IRM
and ARM boxes) is available at http://www.measurementcomputing.com/download.htm

Software for configuring Quicksilver servo motors is available at
http://www.qcontrol.com/software/QuickControl.htm

2. INSTALLATION

	(1) Install Visual Basic 6. As necessary, install Measurement Computing and
		Quicksilver drivers and control software. Example codes for configuring
		motors using the QuickControl software are provided in the directory
		MotorConfigs.
		
	(2) Install the Paleomag Visual Basic code in the desired directory. The default
		directory is C:\Paleomag\Paleomag 2010.
		
		If the C:\Paleomag directory does not exist, after launching the code for
		the first time please edit the Windows Registry to change the location of
		the INI file, as mentioned above. If you install to a different directory,
		you will also need to change the location of the usage log and the help file,
		via the Options dialog box.
		
	(3) Install vbSendMail. Extract the vbSendMail.zip to the target directory
		(e.g., C:\Paleomag\vbSendMail). From the command prompt, register the
		vbSendMail.dll file with Windows by changing to the install directory and
		typing "REGSVR32 vbSendMail.dll" (without quotes).
		
	Optional: If desired, use the procmailrc and Perl scripts provided in the
	MagnetometerAlert directory to set up a Mac as a spoken-word alert system.

3. COMPONENT FILES

FORMS - BACK-END

frmAF - Handles interactions with the AF units
frmDCMotors - Handles interactions with DC motors
frmIRMARM - Handles interactions with IRM and ARM boxes, which occur via frmMCC
frmMCC - Handles interactions with the Measurement Computing card
frmSendMail - Handles email (via the vbSendMail DLL)
frmSquid - Handles communication with SQUID boxes
frmSusceptibilityMeter - Handles communications with Bartington susceptibility bridge
frmVacuum - Handles interactions with vacuum box

FORMS - FRONT-END

frmAbout - The 'About' dialog box
frmCalRod - Calibrate the altitudes
frmChanger - Handles ordering of samples in slots
frmChangerSampOrder - Dialog box for loading samples into a frmChanger
frmDebug - Displays debug messages
frmHelp - HTML viewer for help file
frmLogin - Dialog box for user login
frmMagnetometerControl - Dialog box for automatic and manual sample changer control
frmMeasure - Displays measurements
frmOptions - Dialog box for options
frmProgram - The main program window
frmRerunSamples - Dialog box that 
frmRockmagRoutine - Dialog box for setting AF/Rockmag routine
frmSampleIndexRegistry - Window for loading samples into registry
frmSampleQueueMonitor - Dialog box displaying sample queue
frmSampleSelect - Dialog box for individual sample insertion into frmChanger
frmSettings - Dialog box for settings
frmSplash - Splash screen
frmStats - Displays measurement statistics
frmStepMonitor - Monitors rock mag steps
frmTip - Displays tips on login
frmVRM - Monitors and records changes in moments over time (crude VRM measurements)

MODULES

CBW - Contains Measurement Computing declarations.
modAF - Defines some parameters used by the AF system. Mostly legacy code.
modChanger - Handles movements of sample changer, including sample pickup and drop off
modConfig - Handles loading and saving of configuration parameters from Paleomag.INI
modDataAnalysis - Launches editors/viewers for SAM and sample files
modFlow - Handles flow control (running, paused, and halted)
modMagnetometer - Handles some basic setup routines
modMeasure - Handle measuring and bedding and fold corrections
modMotor.bas - Legacy interface to motor routines
modPrint - Legacy print code (not tested)
modProg - Main program code
modStatusCode - Handles setting of status code levels
modSusceptibility - Handles measuring and correction of susceptibility
modVector3d - Provides vector math

CLASSES

MeasurementBlock - 	A single measurement block (two zeros and four measurements,
					for both samples and handler)
MeasurementBlocks - A set of measurment blocks
RockmagStep - A rock mag step, including description and execution
RockmagSteps - A set of rock mag steps; used in keeping track of steps to be executed
Sample - A sample file, including routines for writing sample data
SampleCommand - A sample command, including routines for execution
SampleCommands - A set of sample commands, as in the sample command queue
SampleIndexRegistration - Representation of a SAM file
SampleIndexRegistrations - A set of sample index registrations, as in the sample index registry
Samples - A set of samples
VectorAngular3D - A vector in spherical coordinates
VectorCartesian3D - A vector in Cartesian coordinates

SUPPORT FILES

MagnetometerAlert - example Perl scripts and procmailrc for setting up a spoken
				    alert system on a Mac with an activated mail server
MotorConfig - example Quickcontrol motor configuration scripts for Quicksilver DC
			  servo motors
vbSendMail.zip - vbSendMail DLL library and documentation
