# Debugging

## The debug log

Turn on **Glf Debug Log** in the VPX options dialog (F6). GLF writes a timestamped file to
`glf_logs/` beside the table.

**Glf Debug Log Level** has two settings:

- **Info** — startup, mode transitions, and whatever your devices log
- **Debug** — every event dispatched, every listener called with its priority, and every player
  variable change

Debug level is verbose enough to slow the table down. Use it to answer a specific question, then
turn it back off.

## Per-device debugging

`Debug = True` on a mode propagates to every device and player in it:

```vbscript
With CreateGlfMode("multiball", 500)
    .Debug = True
End With
```

Or narrow it to one device:

```vbscript
With .Shots("left_ramp")
    .Debug = True
End With
```

This is usually more useful than the global Debug level: you get the internals of the thing you're
investigating without drowning in everything else.

---

## Common problems

### An event never fires

Check, in order:

1. **The name.** Prefixes are inconsistent by device — `balldevice_` for ball devices,
   `auto_fire_coil_` for auto-fire, bare names for shots. See the
   [event reference](event-reference.md#naming-quirks-worth-knowing).
2. **The mode is running.** A listener declared in a mode only exists while it runs. Add a
   temporary `.Add "mode_x_started", Array("debug_marker")` to confirm.
3. **A condition is false.** Debug-level logging shows the listener being skipped.
4. **A shot consumed the switch.** Shots block their switch event for other shots in the same mode.

Debug level logs `<event> has no listeners` for events nothing is listening to — the fastest way to
confirm a typo in either direction.

### A condition never matches

Read `cached-functions.vbs`, written beside the table at startup. It contains the generated
VBScript for every compiled expression. Seeing the actual generated code usually makes the problem
obvious — most often a variable name that doesn't exist, which evaluates to `False`.

Remember undeclared player variables read as `False`, so `{current_player.typo == 0}` is *true*.
Declaring variables with `Glf_SetInitialPlayerVar` avoids the whole class of problem.

### A show doesn't play

- The show must be created **before** the mode that references it. A missing name fails silently.
- Check the `Key` is unique. Two shows sharing a key interfere.
- Check priority: a higher-priority show on the same lights wins the
  [light stack](../concepts/lights-and-shows.md#the-light-stack).

### Lights stay on after a mode ends

Something pushed onto the light stack without popping. Usually a looping show started with a key
that was never stopped. Pair every long-lived `play` with a `stop` on `mode_<name>_stopping`.

### A game won't start

GLF refuses to start until every ball is in the trough. Look for a ball stuck on the playfield or
sitting in a scoop. Check `tnob` matches the number of trough kickers.

### End of ball hangs

Something is holding the `ball_ending` queue and never releasing it — a
[queue relay player](../players/queue-relay-player.md) whose `WaitFor` event never arrives, or a
mode with `UseWaitQueue = True` that never stops. Give such modes a timer as a backstop.

### A ball device keeps re-kicking

`EjectTimeout` is too short — GLF decides the eject failed and fires again while the ball is still
leaving. Raise it. 2000 ms suits most scoops.

### Frame rate drops

Lights are the usual cost. Try **Glf Min Lightmap Update Rate** to throttle lightmap updates, and
look for shows addressing large tags at fast speeds. **Glf Frame Dispatch** caps events processed
per frame; raising it helps a backlog but costs frame time.

---

## The GLF monitor

**Glf Monitor** connects an external monitor application showing live playfield state — switches,
lights, running modes, player variables. Useful for watching what a mode is actually doing without
reading logs.

## MPF config export

With debug logging on, GLF writes an MPF-format config tree to `glf_mpf/` at startup: your modes,
shows, shot profiles, switches, lights and ball devices as YAML.

Its practical use is as a **readable dump of your configuration**. Reading `glf_mpf/modes/<name>/`
is often the quickest way to confirm that a mode is configured the way you think it is, especially
when configuration is spread across several files.

## Reading the running state

Two globals are worth inspecting at a breakpoint or in a debug print:

- `glf_running_modes` — a string listing running modes and priorities
- `glf_BIP` — balls in play

## Resetting persistent state

Delete `<cGameName>_glf.ini` to clear high scores and persisted machine variables.

## See also

[Event reference](event-reference.md) · [Installation](../getting-started/installation.md)
