# Ball device

Anything that catches and holds a ball: the plunger lane, scoops, saucers, VUKs, locks. GLF tracks
what's in each device and coordinates ejects.

```vbscript
With CreateGlfBallDevice("kicker1")
    .BallSwitches = Array("s_Kicker1")
    .EjectTimeout = 2000
    .MechanicalEject = True
    .EjectAllEvents = Array("eject_kicker1")
    .EjectCallback = "Kicker1EjectCallback"
End With
```

Ball devices are machine devices — create them in `ConfigureGlfDevices`, outside any mode.

**Their switches must be in the `glf_switches` collection.**

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `BallSwitches` | — | Array of switches, one per ball position |
| `EjectCallback` | — | Name of a function that fires the coil |
| `EjectTimeout` | 1000 | Ms before an eject is judged failed and retried |
| `EjectEnableTime` | 0 | Ms to hold the coil energised |
| `EntranceCountDelay` | 500 | Ms to settle before counting a ball as arrived |
| `MechanicalEject` | `False` | Ball can leave without the coil — plunger lanes |
| `DefaultDevice` | `False` | This is the plunger; balls from the trough go here |
| `EjectEvents` | — | Events that eject one ball |
| `EjectAllEvents` | — | Events that eject everything |
| `EjectTargets` | — | Switches confirming a ball reached its destination |
| `PlayerControlledEjectEvent` | — | Event for a player-triggered eject |
| `AutoFireOnUnexpectedBall` | `True` | Eject a ball that arrives unexpectedly |
| `ExcludeFromBallSearch` | `False` | Skip during ball search |

## The eject callback

GLF handles when to eject; your callback does the mechanical work:

```vbscript
Sub Kicker1EjectCallback(ball)
    Kicker1.Kick 180, 12
    RandomSoundSaucerKick 1, Kicker1
End Sub
```

The callback receives the ball object, or `Null` when called for the `EjectEnableTime` follow-up
(used for devices needing a hold-then-release).

## The plunger

The plunger lane is a ball device with two distinguishing settings:

```vbscript
With CreateGlfBallDevice("plunger")
    .BallSwitches = Array("s_Plunger1")
    .EjectTimeout = 200
    .MechanicalEject = True
    .DefaultDevice = True
    .EjectCallback = "PlungerEjectCallback"
End With
```

`DefaultDevice = True` makes it the destination for balls served from the trough — every table
needs exactly one. `MechanicalEject = True` tells GLF the ball can leave by the player pulling the
plunger, so a ball that vanishes without a coil fire isn't treated as an error.

## Events posted

| Event | When |
|---|---|
| `balldevice_<name>_ball_entered` | A ball settled in the device |
| `balldevice_<name>_ball_exiting` | A ball is leaving |
| `balldevice_<name>_ejecting_ball` | An eject was commanded |
| `balldevice_<name>_ball_eject_success` | The ball arrived somewhere |

Note the `balldevice_` prefix on events, versus `ball_devices` in the placeholder
`device.ball_devices.<name>.balls`.

```vbscript
With .EventPlayer()
    .Add "balldevice_kicker1_ball_entered", Array("scoop_award", "play_sfx_scoop")
End With
```

## Ejecting

Usually by event:

```vbscript
With CreateGlfBallDevice("kicker1")
    .BallSwitches = Array("s_Kicker1")
    .EjectAllEvents = Array("eject_kicker1")
    .EjectCallback = "Kicker1EjectCallback"
End With
```

```vbscript
With .EventPlayer()
    .Add "scoop_award_complete", Array("eject_kicker1")
End With
```

`EjectEvents` ejects one ball; `EjectAllEvents` empties the device. For a single-ball scoop they're
equivalent.

## Multi-ball devices

A device holding several balls needs one switch per position, in order:

```vbscript
With CreateGlfBallDevice("lock_saucer")
    .BallSwitches = Array("s_Lock1", "s_Lock2", "s_Lock3")
    .EjectTimeout = 2000
    .EjectCallback = "LockEjectCallback"
End With
```

Read the count with `device.ball_devices.lock_saucer.balls`.

## Timeouts and retries

If a ball hasn't left within `EjectTimeout`, GLF fires the coil again. Too short and it
double-kicks a ball that was on its way; too long and a genuinely stuck ball sits there. 2000 ms
suits most scoops; the plunger lane uses a short one (200 ms) because a ball leaving it isn't
coil-driven anyway.

`EntranceCountDelay` is the opposite end — how long to wait before believing a ball has settled, so
a ball bouncing through doesn't register.

## Notes

- Balls are never destroyed; GLF moves the real VPX balls around.
- Devices participate in [ball search](ball-search.md) unless excluded.
- `AutoFireOnUnexpectedBall` means a ball arriving with no reason to be there is kicked back out,
  which quietly recovers from a lot of edge cases.

## See also

[Trough](trough.md) · [Multiball locks](../mode-devices/multiball-locks.md) ·
[Ball hold](../mode-devices/ball-hold.md)
