# Independent ARM bias integration

Source checkpoint: 3 October 2026. Physical instrument acceptance is pending.

Live queue ARM now plans a calibrated axial AF pass and an independent MCC
DC bias voltage. It does not subtract the bias from the alternating field or
apply the IRM incremental-ramp count to ARM. The complete AF lifecycle provides
verified positioning, cancellation and return home. Its immutable treatment
record includes the requested bias and bias setup/cleanup phases.

The circuit follows `VB6/frmIRMARM.frm:SetBiasField`: zero DAC output, wait
0.5 seconds, drive the active-low ARM gate low, wait 0.5 seconds, set the
calibrated voltage, then wait 1 second. Cleanup writes zero, waits 0.5 seconds
and disconnects the gate. The gate-disconnection attempt still runs after a
failed zero-voltage write. Cancellation during setup triggers cleanup. The
queue also clears bias during its general safe-state return, even if motor
communication is disconnected.

Legacy import maps `ARMMax` from gauss to a bias limit in mT, `ARMVoltGauss`
to V/mT, and `ARMVoltMax` to the voltage limit. It leaves the operator's AF peak
and requested bias unchanged. The station profile supplies MCC board 0, DAC
channel 0, active-low gate bit 0, AUXPORT and UNI10VOLTS. Both output mappings
must refer to the same MCC board. Invalid/disabled calibration or an out-of-range
bias rejects the treatment before physical motion or DAC output.

`rapidpy_common.mcc_daq.MccDaq` loads the matching 32/64-bit Universal Library
DLL, probes the configured board without writing outputs, converts voltage
through the vendor calibration function and checks every driver status. The
ctypes interface follows the local `VB6/CBW.BAS` declarations and the vendor
documentation for [cbFromEngUnits](https://files.digilent.com/manuals/Mcculw_WebHelp/Function_Reference/Miscellaneous_Functions/cbFromEngUnits.htm)
and [cbDBitOut](https://files.digilent.com/manuals/Mcculw_WebHelp/Function_Reference/Digital_IO_Functions/cbDBitOut.htm).
Installed vendor drivers and board configuration are required. Drivers are
not installed by this app or copied into the portable bundle.

The 622-test main-app suite passes. Injected ARM/MCC tests cover output ordering,
unmodified alternating field, cancellation, gate cleanup after a DAC failure,
invalid configuration, read-only board probing and rejection of driver errors
before subsequent writes. These tests actuate no instruments. Driver
acknowledgement is not physical field readback or scientific acceptance.

Manual ARM now uses the same queue lifecycle through `ManualArmTreatment`,
requires a specimen identity, and writes immutable attempts plus a SHA-256
artifact index under the configured data folder's `manual_treatments/` directory.
Its worker runs outside the GUI thread. Cancel, Escape and window close request
cancellation and keep the dialog's measurement/changer/AF/IRM leases until cleanup
and evidence publication finish. Invalid inputs or an unwritable evidence
destination fail before actuator output. Evidence publication failures cannot
produce a success report. Queue context and halt callbacks are restored afterward.

Outstanding work includes calibration creation/review controls, physical field/relay verification and
scientific comparison. Pulse IRM/backfield now have queue and manual integration;
see `pulse-irm-development.md` for evidence and remaining qualification. RRM,
motor-configuration migration and other MCC routes remain development tasks.
The previous portable pilot predates this source
checkpoint and must be rebuilt after development stabilizes.
