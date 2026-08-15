# Trough

GLF requires a **physical trough** — balls are moved between kickers rather than created and
destroyed. This is more stable than destroying balls, and it means the ball count is always real.

## Table objects

Add kickers named `swTrough1` … `swTrough<n>`, where *n* is `tnob` (up to 7), plus a kicker named
`Drain`.

`swTrough1` is the release position — the one that kicks balls into the plunger lane. Balls queue
up behind it, and `Drain` is where drained balls land before being fed back.

Model the trough with walls so the balls physically sit where the kickers are; the kickers nudge
balls along the queue.

## Constants

```vbscript
Const tnob = 5      ' total balls the trough holds
Const lob  = 0      ' locked balls not in the trough (captive balls)
Dim gBot            ' ball array, populated by GLF
```

`tnob` determines how many trough kickers GLF creates balls in at startup. It must match the number
of `swTrough` kickers you added.

## What GLF handles

At init, GLF creates the balls and places one in each trough kicker. From then on it:

- Advances balls along the trough as space appears
- Kicks a ball to the plunger lane when the game needs one
- Detects drains and posts `ball_drain`
- Counts balls to decide when a ball, or the game, has ended

You don't write any of that. The trough is the one part of GLF that is entirely built in.

## Events

| Event | When |
|---|---|
| `trough_eject` | The trough kicked a ball toward the plunger |
| `ball_drain` | A ball reached the drain |

`trough_eject` is where ball-release sound and DOF belong:

```vbscript
AddPinEventListener "trough_eject", "on_trough_eject", "OnTroughEject", 2000, Null

Function OnTroughEject(args)
    RandomSoundBallRelease swTrough1
    DOF 110, DOFPulse
End Function
```

`ball_drain` is a **relay event**, so a handler must return the value it was given:

```vbscript
AddPinEventListener GLF_BALL_DRAIN, "ball_drain_sound", "BallDrainSound", 100, Null

Function BallDrainSound(args)
    RandomSoundDrain Drain
    BallDrainSound = args(1)      ' pass the unclaimed ball count along
End Function
```

Returning nothing there breaks the drain chain — the ball count stops propagating and end-of-ball
misbehaves. This is the single most common trough-related mistake.

## Game start check

Before a game can start, GLF confirms every ball is in the trough. If any are missing — stuck on
the playfield, sitting in a scoop — the start request is refused. That's `Glf_BallController`
returning `False` on the `request_to_start_game` relay.

If your table won't start a game, the ball count is the first thing to check.

## Captive balls

Balls that live on the playfield permanently — captive ball lanes — are counted in `lob`, not
`tnob`, and are not created in the trough.

## See also

[Ball device](ball-device.md) · [Ball search](ball-search.md) ·
[Installation](../getting-started/installation.md)
