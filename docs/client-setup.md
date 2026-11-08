# Printing to PrinterOne from other computers

PrinterOne shares a printer that is installed on one Windows PC (the *server*) over the network. Other computers
(the *clients*) send their print jobs to a TCP port on that PC — the same way they would print to a network printer
with a built-in "RAW / JetDirect / port 9100" interface.

## The one rule: the client needs the printer's own driver

PrinterOne does not convert anything. It receives the bytes the client sends and hands them unchanged to the Windows
spooler as a RAW job. So the client must produce data the printer understands:

- **Install the real driver for that printer model on the client** (for example the Canon LBP3300 driver on macOS,
  the HP driver on Windows). A "Generic" or "AirPrint" driver usually sends a language the printer does not speak, and
  the page either does not print or prints garbage.
- If no driver exists for the client's system, that client cannot print through PrinterOne (or any RAW port).

This is the usual cause of "Windows clients print fine but my Mac doesn't" (issue #4): the Mac was set up with a
generic driver.

## Before you start (on the server)

1. Run PrinterOne and note the address shown at start-up (`name -> printer on 0.0.0.0:<port>`) and the server's IP
   address on your network (for example `192.168.1.10`; `ipconfig` shows it).
2. If PrinterOne runs as a service (`printerone service install`), the Windows Firewall rule is already in place.
   If you run `printerone serve` in a console, allow the port in the firewall yourself.
3. If the printer has an allowlist (`allow` in the config), the client's IP must be in it.

Below, `192.168.1.10` and port `9100` are examples — use your server's IP and the printer's port.

## Windows (7, 8, 10, 11)

1. Settings → Printers & scanners → **Add a printer** → *The printer that I want isn't listed*.
2. Choose **Add a printer using an IP address or hostname**.
3. Device type: **TCP/IP Device**; Hostname or IP address: `192.168.1.10`; uncheck *Query the printer*.
4. If asked for more information, choose **Custom** → *Settings…*: Protocol **Raw**, Port number `9100`.
5. Pick the driver for the printer model (Have Disk… if needed) and finish.

Windows XP: Printers and Faxes → Add Printer → Local printer (untick automatic detection) → *Create a new port* →
**Standard TCP/IP Port** → IP `192.168.1.10` → Device type *Custom*: Raw, port `9100` → choose the driver.

## macOS

1. System Settings → Printers & Scanners → **Add Printer…** → the **IP** tab.
2. Address: `192.168.1.10:9100`; Protocol: **HP Jetdirect – Socket**; Queue: leave empty.
3. Use: **Select Software…** and pick the driver for the printer model. Install the manufacturer's macOS driver
   first if it is not in the list. Do not use *AirPrint* or *Generic PostScript* unless the printer itself speaks
   PostScript.

## Linux (CUPS)

```sh
lpadmin -p office -E -v socket://192.168.1.10:9100 -m <ppd-or-driver-name>
```

Or in the CUPS web page (`http://localhost:631`): Administration → Add Printer → **AppSocket/HP JetDirect** →
`socket://192.168.1.10:9100` → choose the driver for the printer model.

## Scripts and devices

Anything that can open a TCP connection can print, as long as it sends the printer's language:

```sh
# PowerShell
$c = New-Object Net.Sockets.TcpClient('192.168.1.10', 9100); $s = $c.GetStream()
$b = [IO.File]::ReadAllBytes('job.prn'); $s.Write($b, 0, $b.Length); $c.Close()
# macOS / Linux
nc 192.168.1.10 9100 < job.prn
```

One connection is one job; close the connection to finish it.

## When it does not print

| Symptom | Check |
|---|---|
| The client cannot connect | The server's firewall; the port and IP; is PrinterOne running (`printerone logs` shows the log file) |
| PrinterOne's log says `rejected: not allowed` | Add the client's IP to that printer's `allow` list |
| Log says `printed` but nothing comes out | The job is in the Windows print queue on the server: open it there; check the printer is online |
| Garbage or blank pages | The client uses the wrong driver (see the rule above) |
| Log says `dropped: idle` | The client stopped sending for 60 s without closing; check the client or the network |

## Note for printers shared from another PC

When PrinterOne runs as a Windows service it uses the LocalSystem account. That account sees printers installed on
the server for all users (USB, LPT, network TCP/IP ports) but not per-user connections to a printer shared by
another PC (`\\otherpc\printer`). Install such printers with a local TCP/IP port instead, or run `printerone serve`
in the user's session.
