# Combo switches

Detects two switches held or hit in relation to one another — most often both flipper buttons at
once, which is the standard input for mid-game menus and action buttons.

```vbscript
With .ComboSwitches("both_flippers")
    .Switch1 = "s_left_flipper"
    .Switch2 = "s_right_flipper"
    .HoldTime = 500
    .EventsWhenBoth = Array("both_flippers_held")
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `Switch1` | — | First switch |
| `Switch2` | — | Second switch |
| `HoldTime` | 0 | Milliseconds both must be held before `EventsWhenBoth` |
| `MaxOffsetTime` | -1 | Maximum ms between the two activating; -1 for no limit |
| `ReleaseTime` | 0 | Milliseconds of release tolerated before the state drops |
| `EventsWhenBoth` | — | Posted when both are active |
| `EventsWhenOne` | — | Posted when exactly one is active |
| `EventsWhenInactive` | — | Posted when both are released |
| `EventsWhenSwitch1` | — | Posted when only switch 1 is active |
| `EventsWhenSwitch2` | — | Posted when only switch 2 is active |

Times accept [placeholders](../concepts/placeholders.md).

Events are posted on *state change*, not repeatedly — moving into "both" posts `EventsWhenBoth`
once, and it won't post again until the state has left and returned.

## Both flippers as an action button

```vbscript
With .ComboSwitches("flipper_combo")
    .Switch1 = "s_left_flipper"
    .Switch2 = "s_right_flipper"
    .HoldTime = 300
    .MaxOffsetTime = 200
    .EventsWhenBoth = Array("action_button_pressed")
    .EventsWhenInactive = Array("action_button_released")
End With
```

`MaxOffsetTime` is what stops normal play triggering this: without it, hitting the left flipper and
then the right two seconds later counts as a combo. With 200 ms, the player has to mean it.

`HoldTime` adds a deliberate delay so a quick alternating flip doesn't register.

## One-sided events

`EventsWhenSwitch1` and `EventsWhenSwitch2` fire when only that switch is active — useful for
menu navigation where each flipper does something different but both together confirms:

```vbscript
With .ComboSwitches("menu_input")
    .Switch1 = "s_left_flipper"
    .Switch2 = "s_right_flipper"
    .EventsWhenSwitch1 = Array("menu_previous")
    .EventsWhenSwitch2 = Array("menu_next")
    .EventsWhenBoth    = Array("menu_select")
    .HoldTime = 200
End With
```

Ordering caveat: pressing left then right passes through the "switch 1 only" state on the way to
"both", so `menu_previous` fires before `menu_select`. Where that matters, drive the single-switch
actions from plain switch events with a condition instead, or accept the extra step and design the
menu around it.

## Release tolerance

`ReleaseTime` keeps the combo alive across a brief release — useful with real cabinet buttons,
where a held press can bounce.

## Notes

- Works with any two switches, not just flippers.
- Flipper *buttons* post `s_left_flipper_active` / `s_right_flipper_active` regardless of whether
  flippers are enabled, so a combo still works when flippers are killed.

## See also

[Timed switches](timed-switches.md) · [Sequence shot](sequence-shot.md)
