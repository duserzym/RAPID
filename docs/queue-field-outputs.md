# Original queue field cutoff evidence

`QueueFieldOutputs` binds the actual native AF ADwin controller, pulse MCC/ADwin
circuit and ARM MCC bias controller into the durable parent queue profile. It
requires the original worker claim, unchanged board/channel calibration and
distinct ARM/pulse outputs before any I/O. This cutoff API never boots a board,
charges/fires a pulse, moves a motor or releases vacuum.

It records a pending `field_outputs` stage before commands, independently zeros
ARM bias and checks its inactive gate with native `cbDBitIn`, then checks/stops
all ten AF processes and acknowledges zero on both ADwin DACs without changing
relays. Pulse recovery independently inhibits fire, zeros charging and enables
bleed, requiring three bounded capacitor readbacks below the accepted discharge
threshold. Shared relay changes require both AF output shutdown and verified
capacitor discharge. A failed stop, DAC command, gate read, capacitor read or relay
readback retains the pending stage and all available component evidence.

Successful cutoff evidence is published as an immutable linked queue event before
returning a typed `QueueFieldOffProof`. It binds the original queue and live circuit
service instance. XY reference/table motion, blank-holder pose, specimen lift phases
and queue pump/valve transitions revalidate this proof before starting their stage
and retain its plain snapshot in the plan. Copied dictionaries, changed calibration,
pending stages and a subsequent AF/ARM/pulse/RRM treatment invalidate authorization.
Existing strict Boolean operator attestations remain compatibility APIs; the automatic
coordinator must use the live typed proof.

An unfinished cutoff can be retried only in original queue recovery, reusing its
stage token and original circuit profile. This does not clear a foreign pending
motion/acquisition stage or finish the parent queue. Original-panel recovery still
must establish specimen support before releasing grip and terminal ownership.

Native DLL acknowledgements prove control commands, process status and digital gate
readback. They do not measure residual magnetic field or analog DAC output voltage.
Capacitor feedback depends on accepted physical calibration. Injected-DLL/serial
regressions cover command ordering, independent cleanup, unsafe evidence, no firing,
publication failure, proof freshness and native transfer borrowing; physical station
qualification remains required. Automatic queue/UI wiring remains open.
