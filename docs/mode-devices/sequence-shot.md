# Sequence shot

Detects a series of switches or events hit in order within a time limit — orbits, ramp-to-ramp
combos, and any "shoot these three in sequence" rule.

```vbscript
With .SequenceShots("left_orbit")
    .SwitchSequence = Array("s_LeftOrb1", "s_LeftOrb2")
    .SequenceTimeout = 500
End With
```

Completing the sequence posts `left_orbit_hit` — the bare device name plus `_hit`, matching a
[shot](shot.md), so downstream rules don't need to care which kind produced it.

```vbscript
With .EventPlayer()
    .Add "left_orbit_hit", Array("score_25000", "play_sfx_orbit")
End With
```

## Configuration

| Property | Meaning |
|---|---|
| `SwitchSequence` | Array of switch names, in order |
| `EventSequence` | Array of events, in order — an alternative to switches |
| `SequenceTimeout` | Milliseconds allowed between steps; 0 means no limit |
| `CancelSwitches` | Switches that abort a sequence in progress |
| `CancelEvents` | Events that abort a sequence in progress |
| `DelaySwitchList` | Switches that pause the sequence clock |
| `DelayEventList` | Events that pause the clock — see below |
| `EnableEvents` / `DisableEvents` | Enable/disable the device |

## Events posted

| Event | When |
|---|---|
| `<name>_hit` | The sequence completed |
| `sequence_shot_<name>_timeout` | A sequence in progress expired |

Note the asymmetry: the completion event uses the bare name, the timeout event carries the
`sequence_shot_` prefix.

## Timeouts

`SequenceTimeout` is the window between consecutive steps, not for the whole sequence. A ball
rolling an orbit takes a few hundred milliseconds between the entry and exit switches; 500 ms is a
typical value. Too generous and unrelated later hits complete the sequence by accident.

## Cancel switches

A ball that enters an orbit but drops into a pop bumper shouldn't complete the orbit when it
eventually rolls over the far switch:

```vbscript
With .SequenceShots("left_orbit")
    .SwitchSequence = Array("s_LeftOrb1", "s_LeftOrb2")
    .SequenceTimeout = 800
    .CancelSwitches = Array("s_Bumper1", "s_Bumper2", "s_Bumper3")
End With
```

## Delaying the clock

`DelayEventList` accepts an event with a `:ms` suffix giving how long to suspend the timeout:

```vbscript
.DelayEventList = Array("ball_in_saucer:2000")
```

While delayed, the sequence clock doesn't advance — for a ball held by a scoop or magnet part-way
through a combo. This `event:ms` form is specific to this property.

## Event sequences

Sequences can be built from game events instead of raw switches, which lets a combo be defined
over rules rather than hardware:

```vbscript
With .SequenceShots("ramp_combo")
    .EventSequence = Array("left_ramp_hit", "right_ramp_hit", "left_ramp_hit")
    .SequenceTimeout = 4000
End With
```

## Multiple concurrent sequences

GLF tracks several in-flight sequences at once, so during multiball two balls can each be part-way
through the same sequence independently.

## See also

[Shot](shot.md) · [Combo switches](combo-switches.md)
