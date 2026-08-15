# Modes

A mode is a bundle of rules with a lifetime. While it runs, its listeners are registered and its
devices are live; when it stops, all of that is torn down. Modes are how GLF keeps a large ruleset
from turning into a tangle of global state.

```vbscript
With CreateGlfMode("multiball", 500)
    .StartEvents = Array("start_multiball")
    .StopEvents  = Array("multiball_multiball_ended", "ball_will_end")
    ' … devices and players …
End With
```

`CreateGlfMode(name, priority)` returns the mode, creating it if needed. Calling it again with the
same name returns the existing mode rather than a second one, so you can add to a mode from
another file.

---

## Priority

Priority does three things:

1. **Orders event handlers.** Higher priority handlers run first.
2. **Sets light show priority.** Shows a mode plays are pushed onto each light's stack at the
   mode's priority; the highest wins the light. This is how a multiball mode's lighting overrides
   the base mode's without either knowing about the other.
3. **Provides a base for offsets.** A listener written `event.10` runs at mode priority + 10.

Rough conventions from working tables:

| Range | Typical use |
|---|---|
| 10–100 | Attract, background/ambient |
| 100–200 | Base mode, always-on scoring |
| 200–600 | Feature modes, multiballs |
| 1000+ | Bonus, high score, tilt — things that must sit on top |

Leave gaps. Retro-fitting a mode between two adjacent numbers is tedious.

---

## Start and stop events

```vbscript
.StartEvents = Array("ball_started")
.StopEvents  = Array("ball_will_end", "tilt")
```

Both accept the full [event syntax](events.md#event-string-syntax), so modes can be conditional:

```vbscript
.StartEvents = Array("ball_started{current_player.ball == 3}")
```

Starting a mode posts, in order:

1. `mode_<name>_starting` — a queue event; devices activate here
2. `mode_<name>_started` — the one you hook

Stopping posts `mode_<name>_stopping` then `mode_<name>_stopped`.

Hook `mode_<name>_started` for anything that should happen when the mode begins, and
`mode_<name>_stopping` for cleanup — the latter fires while the mode's devices are still alive.

```vbscript
With .EventPlayer()
    .Add "mode_base_started",  Array("stop_attract_mode", "play_mus_ambient_loop")
    .Add "mode_base_stopping", Array("stop_mus_ambient_loop")
End With
```

Stopping a mode that isn't running does nothing, so redundant stop events are harmless.

---

## Two useful flags

### `UseWaitQueue`

```vbscript
.UseWaitQueue = True
```

When this mode is started by a **queue event**, the queue is held open until the mode stops. That
is how a bonus-tally or high-score mode delays end-of-ball: start it on `ball_ending` with
`UseWaitQueue = True`, and the rest of the ball-end sequence waits.

### `GameMode`

```vbscript
.GameMode = False
```

Marks a mode as running outside normal ball play. The built-in high score mode uses it, since it
runs during `game_ending` when there's no ball in play.

---

## What lives in a mode

Two families, both accessed as properties of the mode.

**Mode devices** — stateful things, each with a name:

```vbscript
With .Shots("left_ramp")
With .Timers("countdown")
With .Multiballs("main")
With .BallSaves("new_ball")
With .StateMachines("progress")
```

See the [mode devices reference](../mode-devices/).

**Players** — map events to actions. Most are singletons on the mode:

```vbscript
With .EventPlayer()
With .VariablePlayer()
With .ShowPlayer()
With .SegmentDisplayPlayer()
```

See the [players reference](../players/).

Both are created lazily on first access, so mentioning `.Timers("foo")` twice gives you the same
timer — handy for configuring one across several `With` blocks.

---

## Enable and disable

Most mode devices accept `EnableEvents` and `DisableEvents`, giving a second level of control
inside a running mode:

```vbscript
With .Shots("jackpot")
    .Profile = "jackpot_profile"
    .Switch = "s_Scoop"
    .EnableEvents  = Array("jackpot_lit")
    .DisableEvents = Array("jackpot_collected")
End With
```

The distinction: **activation** follows the mode's lifetime and isn't optional; **enable/disable**
is yours to drive. A device with no `EnableEvents` generally enables itself when the mode starts,
which is what you want most of the time.

---

## Debugging a mode

```vbscript
.Debug = True
```

Propagates to every device and player in the mode, writing their internals to the debug log. Set
it on one mode at a time — it's verbose. See [Debugging](../appendix/debugging.md).

---

## A worked example

```vbscript
Sub CreateMultiballMode()

    With CreateGlfMode("multiball", 500)
        .StartEvents = Array("multiball_main_started")
        .StopEvents  = Array("multiball_main_ended", "ball_will_end")

        With .EventPlayer()
            .Add "mode_multiball_started", Array("play_mus_multiball_loop", "stop_mus_ambient_loop")
            .Add "mode_multiball_stopping", Array("stop_mus_multiball_loop", "play_mus_ambient_loop")
            .Add "s_Scoop_active", Array("award_jackpot")
        End With

        With .VariablePlayer()
            With .EventName("award_jackpot")
                With .Variable("score")
                    .Action = "add"
                    .Int = "{current_player.jackpot_value}"
                End With
            End With
        End With

        With .ShowPlayer()
            With .EventName("mode_multiball_started")
                .Key = "key_mb_show"
                .Show = "multiball_running"
                .Loops = -1
                With .Tokens()
                    .Add "color", MultiballColor
                End With
            End With
            With .EventName("mode_multiball_stopping")
                .Key = "key_mb_show"
                .Action = "stop"
            End With
        End With

    End With

End Sub
```

Note the show pairing: the same `Key` is used to start and stop it. Keys are how you address a
running show later — see [Lights & shows](lights-and-shows.md).
