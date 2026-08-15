# Segment display player

Writes text to [segment displays](../devices/segment-display.md) — alphanumeric or numeric
displays built from lights.

```vbscript
With .SegmentDisplayPlayer()
    With .EventName("mode_base_started")
        With .Display("player1")
            .Text = "{players[0].score:0>2,}"
        End With
        With .Display("ball")
            .Text = "{current_player.ball:0>2}"
        End With
    End With
End With
```

Structure: an `EventName` block per triggering event, containing a `Display` block per display to
update. The display name is the one given to `GlfLightSegmentDisplay`.

## Properties

| Property | Default | Meaning |
|---|---|---|
| `Text` | — | Text to show; supports [placeholders](../concepts/placeholders.md) and formatting |
| `Key` | — | Handle for this entry, so it can be removed later |
| `Action` | `"add"` | `add` or `remove` |
| `Priority` | 0 | Added to the mode's priority |
| `Flashing` | `"off"` | `off`, `all`, or `mask` |
| `FlashMask` | — | With `Flashing = "mask"`, which characters flash |
| `Color` | white | Segment colour |
| `Expire` | 0 | Auto-remove after N seconds |
| `Transition` | — | Animated entry — see below |
| `TransitionOut` | — | Animated exit when this entry is replaced |

## Live values

Text is re-evaluated when the placeholder's underlying variable changes, so a score display set
once at mode start keeps up on its own:

```vbscript
With .Display("player1")
    .Text = "{players[0].score:0>2,}"
End With
```

Format specs matter here. `{players[0].score:0>2,}` means zero-padded, right-aligned, minimum width
2, with thousands separators. See [Placeholders](../concepts/placeholders.md#formatting).

## The display stack

Entries stack per display by priority, like lights. A higher-priority entry takes over; remove it
and the previous text returns. That's how a transient message overlays a score without anything
having to remember what the score display said:

```vbscript
With .EventName("jackpot_awarded")
    With .Display("all")
        .Key = "jackpot_msg"
        .Text = "JACKPOT"
        .Priority = 500
        .Expire = 2
    End With
End With
```

With `Expire = 2` the entry removes itself after two seconds. Without it, remove explicitly:

```vbscript
With .EventName("stop_flash_player1_score")
    With .Display("player1")
        .Key = "p1_score_flash"
        .Action = "remove"
    End With
End With
```

Key and priority together are what make this work — reuse the same `Key` when you add and remove.

## Flashing

```vbscript
With .EventName("flash_player1_score")
    With .Display("player1")
        .Key = "p1_score_flash"
        .Text = "{players[0].score:0>2}"
        .Flashing = "all"
        .Priority = 100
    End With
End With
```

`Flashing = "all"` blinks the whole display — the standard "this player is up" indicator.
`"mask"` blinks only positions marked in `FlashMask`, one character per position, `F` for flashing.

## Transitions

```vbscript
With .Display("all")
    .Text = "EXTRA BALL"
    With .Transition()
        .TransitionType = "push"     ' "push" or "cover"
        .Direction = "right"         ' "right" or "left"
    End With
End With
```

`push` slides the old text out as the new arrives; `cover` draws the new text over it. Transition
smoothness is set per display by `DefaultTransitionUpdateHz`.

`TransitionOut` animates when *this* entry is replaced by a different key.

## Notes

- Text longer than the display is truncated, not scrolled.
- Text entries belonging to a mode are removed when it stops.
- Multi-display groups let you write one string across several physical displays — see
  [Segment display](../devices/segment-display.md).

## See also

[Segment display](../devices/segment-display.md) · [Placeholders](../concepts/placeholders.md)
