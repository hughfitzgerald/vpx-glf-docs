# Ball search

Pulses coils to dislodge a ball that has gone missing — stuck behind a post, wedged in a ramp, or
sitting somewhere with no switch.

```vbscript
With EnableGlfBallSearch()
    .Timeout = 15000
    .SearchInterval = 300
    .BallSearchWaitAfterIteration = 5000
End With
```

Call `EnableGlfBallSearch()` from `ConfigureGlfDevices`. It's off unless you enable it.

| Property | Default | Meaning |
|---|---|---|
| `Timeout` | 15000 | Ms of inactivity before a search starts |
| `SearchInterval` | 150 | Ms between pulsing successive devices |
| `BallSearchWaitAfterIteration` | 10000 | Ms to wait before searching again |

## When it runs

A search starts when all of these hold:

- A game is in progress with at least one ball in play
- No switch has been hit for `Timeout` ms
- The plunger lane doesn't hold the ball
- Balls held by a [ball hold](../mode-devices/ball-hold.md) are discounted

Any switch hit resets the timer.

## Flipper cradling

Enabling ball search also creates a [timed switch](../mode-devices/timed-switches.md) watching the
flipper buttons. Holding a flipper for three seconds posts `flipper_cradle`, which suspends the
search; releasing posts `flipper_release` and resumes it.

Without this, a player cradling the ball to read the display would trigger a search — the machine
deciding the ball is lost while they're holding it.

## What gets pulsed

GLF works through devices in phases, pulsing coils and briefly activating diverters and drop
targets. Devices participate unless excluded:

```vbscript
With CreateGlfBallDevice("kicker1")
    .BallSwitches = Array("s_Kicker1")
    .ExcludeFromBallSearch = True
    .EjectCallback = "Kicker1EjectCallback"
End With
```

`ExcludeFromBallSearch` is available on ball devices, auto-fire devices, diverters and drop
targets. Exclude anything that would be pointless (a device that can't hold a stuck ball) or
harmful (a coil that would fire a held ball into play).

The plunger lane is skipped automatically.

## Events posted

| Event | When |
|---|---|
| `ball_search_started` | A search began |
| `ball_search_stopped` | A search ended |

Worth hooking for a display message so the player understands what the machine is doing:

```vbscript
With .EventPlayer()
    .Add "ball_search_started", Array("show_ball_search_message")
    .Add "ball_search_stopped", Array("clear_ball_search_message")
End With
```

## Tuning

`Timeout` is a judgement call. Too short and it fires during ordinary lulls — a slow ramp shot, a
player thinking. Too long and a genuinely stuck ball leaves the game hanging. 15–20 seconds is
typical, and the flipper-cradle suspension covers the common false positive.

## See also

[Ball device](ball-device.md) · [Timed switches](../mode-devices/timed-switches.md)
