# Email notifications in `rapid_main`

RAPID emails people when a measurement run needs them. The implementation is
ported from the email path of
[ASC_oven_control](https://github.com/Institute-for-Rock-Magnetism/ASC_oven_control)
(`infrastructure/notify.py`) and follows VB6 `frmSendMail.MailNotification`.

## Setup

1. **File → Email Notifications…**
2. Tick **Send email notices**, then add **Always copy** addresses: the lab
   manager and/or a status-monitor inbox.
3. Choose how mail leaves the PC:
   * **SMTP server:** host, port, security, username and password. The password
     is stored **encrypted with Windows DPAPI**, readable only by this Windows
     user on this PC.
   * **Classic Outlook on this PC:** no password is stored. Outlook must be
     signed in to an account.
4. **Send Test Email**, then **Save**.

To include an operator, enter their email at **login** (File → Log Out). They
then receive this session's notices, like VB6 `LoginEmail`.

**Import from VB6 INI…** copies the VB6 `[Email]` host, port, SSL, sender and
CC/status-monitor list. It never copies the password. VB6 kept that password
in the INI in plain text, so enter it here instead.

## What is sent

Each notice carries a VB6 status code. You choose which codes are sent.

| Code | When | Default |
|---|---|---|
| Red — emergency | Queue run ended with an error; manual measurement error | on |
| Orange — attention | Queue run halted; manual measurement stopped; "Sample done. Please remove sample." | on |
| Yellow — oops | Recoverable issues (reserved for re-measure / flux-jump notices) | off |
| Green — normal | Queue run complete | on |

The body starts with the operator. The footer lists sample, step, station, time
and code, as in VB6.

## Guarantees

* Sending never blocks or interrupts a measurement. Each notice goes out on a
  background thread with a timeout and retries.
* Failures are recorded and never raised.
* Every attempt is logged to `notifications.log`, next to the RAPID config.
  Passwords are never logged.
* Settings are stored in `notifications.json` next to the RAPID config
  (`%USERPROFILE%\.rapid\` by default). They are never in the repository, and
  they are kept out of the measurement configuration fingerprint, so editing
  recipients mid-run cannot change an acquisition's audit context.

## Implementation

| Piece | Location |
|---|---|
| Settings, DPAPI secrets, VB6 import, message composition, SMTP/Outlook delivery | `RapidPy/rapidpy_common/notify.py` |
| Settings dialog | `RapidPy/rapid_main/rapid_main/dialogs/notifications.py` |
| Hooks: terminal queue state, manual sample finish, operator email at login | `RapidPy/rapid_main/rapid_main/app.py` (`notify`, `_notify_queue_terminal`, `_notify_manual_sample_finished`) |
| Tests | `RapidPy/rapid_main/tests/test_notifications.py` |
