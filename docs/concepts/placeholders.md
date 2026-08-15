# Placeholders & dynamic values

Almost any value in a GLF config can be a live expression instead of a constant. Wrap it in braces:

```vbscript
.Int = 5000                                  ' constant
.Int = "{current_player.spinner_value}"      ' evaluated when the event fires
.Int = "{current_player.base * current_player.multiplier}"
```

The same syntax powers event conditions (`event{...}`) and segment display text.

Expressions are compiled to VBScript functions once during `Glf_Init` and cached, so evaluation at
runtime is a function call, not a parse. This is also why `ConfigureGlfDevices` must run before
`Glf_Init`.

---

## Available sources

### `current_player.<var>`

A player variable for whoever is up.

```vbscript
.Int = "{current_player.bonus_multiplier}"
.Add "s_Ramp_active{current_player.ramps_made >= 3}", Array("ramp_complete")
```

Built-ins available on every player: `score`, `ball`, `extra_balls`, `initials`, `number`, plus
`tilt_warnings` once a tilt device has run. Your own variables come from
`Glf_SetInitialPlayerVar` — see [Variables](variables.md).

A variable that has never been set reads as `False`, which compares equal to `0` in VBScript. In
practice that means `{current_player.never_set == 0}` is true, which is usually what you want, but
it also means typos fail quietly rather than loudly.

### `players[n].<var>`

A specific player, zero-indexed — `players[0]` is player 1. Mostly for score displays:

```vbscript
With .Display("player1")
    .Text = "{players[0].score:0>2,}"
End With
```

### `machine.<var>`

A machine variable — shared across players, optionally persisted to disk.

```vbscript
.Add "s_Lock_active{machine.bottom_ball_locked == 0}", Array("lock_ball")
```

### `device.<type>.<name>.<attribute>`

Live state from a device. The available types and attributes:

| Placeholder | Returns |
|---|---|
| `device.timers.<name>.ticks` | Current timer value |
| `device.timers.<name>.ticks_remaining` | Ticks left before the end value |
| `device.ball_devices.<name>.balls` | Balls currently in the device |
| `device.state_machines.<name>.state` | Current state name |
| `device.shot_groups.<name>.common_state` | Shared state index, or empty if shots differ |
| `device.multiballs.<name>.enabled` | Boolean |
| `device.multiball_locks.<name>.enabled` | Boolean |
| `device.multiball_locks.<name>.locked_balls` | Count |
| `device.ball_holds.<name>.balls_held` | Count |
| `device.diverters.<name>.enabled` / `.active` | Boolean |
| `device.sound_buses.<name>.volume` / `.simultaneous_sounds` | Bus settings |

Typical use — driving a countdown display straight from a timer:

```vbscript
With .EventName("timer_bonus_countdown_tick")
    With .Display("all")
        .Text = "BONUS {device.timers.bonus_countdown.ticks}"
    End With
End With
```

### `modes.<name>.active`

Whether a mode is running. The mode name here is the bare name, without the `mode_` prefix:

```vbscript
.Add "s_Ramp_active{modes.multiball.active == True}", Array("ramp_jackpot")
```

### `game.<attribute>`

| Placeholder | Returns |
|---|---|
| `game.tilted` | Boolean |
| `game.balls_per_game` | 3 or 5, per the table option |
| `game.balls_in_play` | Live ball count |

### `kwargs.<key>`

Data attached to the event currently being handled:

```vbscript
With .EventName("player_added{kwargs.num == 3}")
    With .Display("player3")
        .Text = "{players[2].score:0>2,}"
    End With
End With
```

---

## Arithmetic and conditionals

Placeholder contents are VBScript expressions after substitution, so arithmetic works:

```vbscript
.Int = "{current_player.jackpot_base * current_player.multiplier}"
.Int = "{device.timers.hurry_up.ticks * 10000}"
```

There is also a Python-style conditional form, using ` if ` / ` else ` **with surrounding spaces**:

```vbscript
.Int = "{50000 if current_player.mode_level > 2 else 10000}"
```

---

## Formatting

Add a format spec after a colon. Only meaningful in text output — segment displays and slides.

```
{value:<pad><align><width>}
```

| Piece | Meaning |
|---|---|
| pad char | Character to fill with — usually `0` or a space |
| align | `>` right, `<` left, `^` centre |
| width | Total field width |
| `,` | Insert thousands separators (place it anywhere in the spec) |

```vbscript
.Text = "{players[0].score:0>2,}"      ' 1,234,500 padded to at least 2 chars
.Text = "{current_player.ball:0>2}"    ' 03
.Text = "{current_player.name: ^16}"   ' centred in a 16-character display
```

A boolean `False` formats as an empty string, which is a convenient way to blank a display when a
variable is unset.

---

## Where placeholders work

Broadly: anywhere a value is stored as a *dynamic input* rather than a plain property. That covers
variable player values, timer start/end values and control values, ball save times, multiball ball
counts, tilt settings, combo switch times, sequence timeouts, and segment display text.

Things that are **not** dynamic and need literal values: light and switch names, show names, shot
profile names, mode names and priorities, and token keys. If a placeholder appears to be ignored,
this is usually why.

---

## Debugging expressions

Turn on *Glf Debug Log* with level Debug (see [Debugging](../appendix/debugging.md)). GLF writes
the generated VBScript for every compiled expression to `cached-functions.vbs` beside the table —
reading it is by far the fastest way to see why a condition isn't matching.
