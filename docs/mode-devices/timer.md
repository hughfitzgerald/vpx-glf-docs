# Timer

Counts up or down, posting events as it ticks. Used for mode countdowns, hurry-ups, cooldowns and
anything that needs a repeating beat.

```vbscript
With .Timers("mode_countdown")
    .StartValue = 30
    .EndValue = 0
    .Direction = "down"
    .TickInterval = 1000
    .StartRunning = True
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `StartValue` | 0 | Value at start or reset |
| `EndValue` | -1 | Value at which the timer completes; -1 counting up means never |
| `Direction` | `"up"` | `"up"` or `"down"` |
| `TickInterval` | 1000 | Milliseconds per tick |
| `MaxValue` | — | Ceiling for `add` and `jump` |
| `StartRunning` | `False` | Start as soon as the mode starts |
| `RestartOnComplete` | `False` | Automatically restart on completion |
| `ControlEvents` | — | Event-driven actions — see below |

`StartValue`, `EndValue` and control values accept
[placeholders](../concepts/placeholders.md), so a countdown can shorten as difficulty rises:

```vbscript
.StartValue = "{30 - (current_player.mode_level * 5)}"
```

## Events posted

| Event | Kwargs |
|---|---|
| `timer_<name>_started` | `ticks`, `ticks_remaining` |
| `timer_<name>_tick` | `ticks`, `ticks_remaining` |
| `timer_<name>_complete` | `ticks`, `ticks_remaining` |
| `timer_<name>_stopped` | `ticks`, `ticks_remaining` |
| `timer_<name>_paused` | `ticks`, `ticks_remaining` |
| `timer_<name>_time_added` | `ticks`, `ticks_added`, `ticks_remaining` |
| `timer_<name>_time_subtracted` | `ticks`, `ticks_subtracted`, `ticks_remaining` |

Current value is also readable as `device.timers.<name>.ticks` and `.ticks_remaining`.

## Control events

Actions are driven by events. Each `With .ControlEvents()` adds one entry:

```vbscript
With .Timers("hurry_up")
    .StartValue = 20
    .EndValue = 0
    .Direction = "down"
    .TickInterval = 500
    With .ControlEvents()
        .EventName = "hurry_up_started"
        .Action = "start"
    End With
    With .ControlEvents()
        .EventName = "jackpot_collected"
        .Action = "stop"
    End With
    With .ControlEvents()
        .EventName = "ramp_made"
        .Action = "add"
        .Value = 5
    End With
End With
```

| Action | Effect |
|---|---|
| `start` | Start running |
| `stop` | Stop; value is kept |
| `reset` | Return to `StartValue` |
| `restart` | Reset and start |
| `pause` | Pause; with a `Value` in ms, auto-resume after that long |
| `add` | Add `Value` ticks |
| `subtract` | Subtract `Value` ticks |
| `jump` | Set the value to `Value` |
| `set_tick_interval` | Set interval to `Value` ms |
| `change_tick_interval` | Multiply interval by `Value` |
| `reset_tick_interval` | Return to the configured interval |

`EventName` takes conditions and priority offsets like any listener:

```vbscript
With .ControlEvents()
    .EventName = "spinner_hit{current_player.spinner_value == 1000}"
    .Action = "restart"
End With
```

## Displaying a countdown

```vbscript
With .SegmentDisplayPlayer()
    With .EventName("timer_mode_countdown_tick")
        With .Display("all")
            .Key = "countdown"
            .Text = "TIME {device.timers.mode_countdown.ticks:0>2}"
            .Priority = 200
        End With
    End With
End With
```

## Scoring from remaining time

A hurry-up award that decays is just a placeholder over the timer:

```vbscript
With .VariablePlayer()
    With .EventName("hurry_up_collected")
        With .Variable("score")
            .Action = "add"
            .Int = "{device.timers.hurry_up.ticks * 10000}"
        End With
    End With
End With
```

## Speeding up

`change_tick_interval` multiplies the interval, so values below 1 accelerate:

```vbscript
With .ControlEvents()
    .EventName = "timer_countdown_tick{device.timers.countdown.ticks < 5}"
    .Action = "change_tick_interval"
    .Value = 0.8
End With
```

## Notes

- Counting up with `EndValue = -1` runs indefinitely — fine for an elapsed-time counter.
- Counting down completes when the value reaches or passes `EndValue`, which defaults to -1; set
  it to 0 explicitly for a normal countdown.
- Timers stop and are cleaned up when their mode stops.
- Timer names are global for placeholder lookup, so keep them unique across modes.

## See also

[Counter](counter.md) · [Placeholders](../concepts/placeholders.md)
