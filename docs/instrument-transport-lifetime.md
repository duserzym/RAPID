# SQUID and susceptibility transport lifetime

The shared `updown_control` RawSquidClient and the main app's susceptibility
serial client now retain their exact original serial handle when close raises
or returns while the port remains open. A subsequent close retries that handle;
it cannot silently discard an unsettled connection. The susceptibility client
also settles a retained disconnected handle before opening a replacement.

SquidBackendAdapter retains its original reader and baseline until disconnect
succeeds. Its connected state reflects the reader's actual connection rather
than the presence of a reader object. `prepare_raw_client` creates that reader
without opening a port, resetting counters or taking readings.
`connect_for_acquisition` opens that same reader without the diagnostic baseline
read. The diagnostic connection test still takes its baseline explicitly.

Native queue preflight can now compose bracketed acquisition using the prepared
raw client while the instrument remains disconnected. Native SQUID acquisition
requires the exact original pending acquisition stage/token before connection;
the bracketed service then owns reset/latch/read operations. Injected tests use
the real adapter, reader and composition path and inspect the durable stage at
the connect call. An absent valid holder still fails scientific acquisition and
retains its pending owner; the test does not fabricate a qualified specimen block.

Original native queue adapter/client/serial identity and port configuration now
bind to QueueInstrumentLifetime; terminal independent close evidence precedes root
publication and MainWindow release. See queue-instrument-settlement.md. Live/restart
recovery controls, failed-open cleanup, and physical/scientific acceptance remain.
The separate VRM logger serial client is not changed by this checkpoint.
No physical instrument was actuated by the injected tests.
