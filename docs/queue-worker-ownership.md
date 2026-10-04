# Native queue worker ownership and measured backend binding

Both QueueCommandWorker and MeasurementWorker enter a declared backend
`queue_worker_claim` on the worker thread. QueueHardwareBackend borrows the original
QueueStageStore session and retains its claim through stage I/O, recovery attempts,
artifact publication and cancellation-hook cleanup. Completion is emitted after
claim release; existing GUI settlement still waits for the actual QThread to exit.
Unbound backends retain their existing worker contract.

An already claimed/released session or mismatched geometry/store prevents instrument
I/O. A rejected competing command worker cannot clear the current owner's halt hook.
MeasurementWorker reports claim/hook failure once as an aborted result. The parent
queue journal remains latched after a worker exits; claim release is not terminal
output or workflow release.

`bind_queue_coordinator` attaches the native backend to its original routed motors,
scientific stage profiles, AF/ARM/pulse circuit instances and fixed journal under
the existing queue claim. Load/return revalidate these bindings before commands.
Transfer cancellation includes the backend's cooperative halt probe. Loading calls
the composed transfer coordinator and binds its measured specimen geometry before
acquisition. Return clears geometry only after verified original-slot release and
rod clearance; failed return retains the binding and grip. SQUID transport recovery
records survive geometry/service retirement for the final worker artifacts.

Borrowed queue cleanup cannot fall through to generic unstaged motor commands.
An unfinished stage requires original queue recovery; a verified loaded specimen
returns through its original coordinator. Missing coordinator, unknown rod state
and blank-holder cleanup fail closed pending their explicit integration.

Injected native DLL/serial tests cover QThread acquisition/publication ownership,
competing claims, failure retention, released/mismatched sessions, physical service
binding, scientific-setting changes, native measured load/return, cancellation and
transport evidence retention. Ordinary measurement/command-worker regression tests
also verify their prior contracts.

MainWindow still must create the durable queue and native coordinator at startup,
hold all participating device leases, supply each measurement's original slot/file,
replace legacy blank/InitUp/Flip commands and perform original-panel terminal
recovery/release. This worker/backend integration does not establish physical station
qualification or complete the full automatic UI path.
