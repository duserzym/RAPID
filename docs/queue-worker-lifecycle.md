# Queue worker ownership and cancellation

Native `QueueHardwareBackend` initialization, holder acquisition, changer movement
and flip commands run in a Qt worker. Their command leases and the queue's changer
lease remain held until the thread has stopped, including recovery after failures
or cancellation. Halt sets a cooperative cancellation event; it does not issue
competing motor I/O from the GUI thread. Pause prevents subsequent commands without
interrupting an in-flight positioning command. Resume cannot overlap that command.

Normal completion and cancellation run native terminal recovery in a worker too.
Recovery failures produce an error terminal state. Closing the main window waits
for active acquisition, command workers and recovery. Owned timers perform queued
advancement and terminal checks without callbacks outliving the window.
The DC Motors close timer likewise waits for its terminal callback to clear the
command reference; a stopped thread alone cannot dismiss the dialog prematurely.

Acquisition's completion signal does not release measurement ownership or advance
the queue while its worker is still running. Error notifications retain the same
ownership until terminal cleanup. Queue cancellation waits for acquisition to exit
before requesting terminal queue recovery. Native recovery runs directly in the
measurement worker, so an outer halt/timeout wrapper cannot misclassify a completed
recovery solely because the operator pressed Halt.

The native halt probe is installed in motor motion and bracketed acquisition.
Settling checks cancellation in 50 ms intervals. Motion and transport phases check
before I/O; cancellation during the final SQUID read prevents returning a completed
block. Cancellation is an `InterruptedError`, not a transport fault eligible for
whole-block read retry. Transport and motion calls still use their native bounds.

Regression tests inject bounded transports and controlled threads to check GUI
responsiveness, premature cancellation without I/O, late cancellation, independent
terminal signaling, pause/resume exclusion, cleanup failure, acquisition handoff
and lease retention. No physical instruments are actuated.

This checkpoint does not qualify live queue operation. Queue vacuum must gain a
coordinated lifetime: the main vacuum panel's diagnostic hold currently prevents
native queue preflight, while a disconnected or released vacuum prevents queue
startup. Durable ordinary acquisition and non-treatment motion recovery also
remain required. Preserve the legacy sample-transfer order when implementing that
coordinator; physical gripper, pressure, clearance and scientific acceptance remain
open. The portable pilot must be rebuilt after the remaining software work.
