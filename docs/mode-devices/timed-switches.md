# Timed switches

Posts events when a switch is held for a set time — the standard way to detect a cradled ball or a
long button press.

```vbscript
With .TimedSwitches("cancel_game")
    .Switches = Array("s_start")
    .Time = 2000
    .EventsWhenActive = Array("glf_game_cancel")
End With
```

Holding the start button for two seconds cancels the game. `glf_game_cancel` is a built-in event
GLF acts on.

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `Switches` | — | Array of switches to watch |
| `Time` | 0 | Milliseconds a switch must be held |
| `EventsWhenActive` | — | Posted once a switch has been held that long |
| `EventsWhenReleased` | — | Posted when all watched switches are released |

`Time` accepts a [placeholder](../concepts/placeholders.md).

## How several switches interact

The switch list is treated as a group: `EventsWhenActive` fires when the *first* switch reaches the
hold time, and `EventsWhenReleased` only when the *last* one is released. Holding a second switch
while the first is already active doesn't post again.

That's what makes the flipper-cradle case work with one device:

```vbscript
With .TimedSwitches("flipper_cradle")
    .Switches = Array("s_left_flipper", "s_right_flipper")
    .Time = 3000
    .EventsWhenActive = Array("flipper_cradle")
    .EventsWhenReleased = Array("flipper_release")
End With
```

Either flipper held for three seconds posts `flipper_cradle`; it clears when both are let go. This
is exactly how [ball search](../devices/ball-search.md) knows the player is holding a ball and
suspends itself.

## Ball held in a lane

```vbscript
With .TimedSwitches("ball_parked")
    .Switches = Array("s_Kicker1")
    .Time = 5000
    .EventsWhenActive = Array("ball_stuck_in_kicker")
End With
```

## Notes

- The timer restarts from scratch on each activation — a switch that opens and closes doesn't
  accumulate held time.
- Releasing before `Time` posts nothing at all, not even `EventsWhenReleased`.
- For "two switches at once" rather than "one held a while", use
  [combo switches](combo-switches.md).

## See also

[Combo switches](combo-switches.md) · [Ball search](../devices/ball-search.md)
