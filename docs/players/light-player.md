# Light player

Sets lights to a colour on an event. For static colour changes; use the
[show player](show-player.md) when you need a sequence.

```vbscript
With .LightPlayer()
    With .EventName("mode_base_started")
        With .Lights("GI")
            .Color = GIColor3000k
            .Fade = 300
        End With
    End With
End With
```

| Property | Default | Meaning |
|---|---|---|
| `Color` | `"ffffff"` | Six-digit hex, no `#` |
| `Fade` | none | Fade duration in ms |
| `Priority` | 0 | Added to the mode's priority |

The `Lights` key is a light name or a [tag](../concepts/lights-and-shows.md#tags). Tags expand at
startup, so `"GI"` addressing forty lights costs nothing extra to write.

## Several lights per event

```vbscript
With .EventName("mode_multiball_started")
    With .Lights("GI")
        .Color = "ffffff"
        .Fade = 200
    End With
    With .Lights("L20")
        .Color = JackpotColor
    End With
End With
```

## Lifetime

Light player entries push onto the [light stack](../concepts/lights-and-shows.md#the-light-stack)
at the mode's priority, and are popped automatically when the mode stops — the lights revert to
whatever the next mode down had set. You don't need a matching "turn it off" entry.

## Light player or show player?

| | Light player | Show player |
|---|---|---|
| Steps | One | Many |
| Cost | Cheap — one stack push | Per-step scheduling |
| Cleanup | Automatic on mode stop | Automatic on mode stop, or explicit `stop` by key |
| Tokens | No | Yes |

For "GI goes amber while this mode runs", the light player. For anything that flashes, chases or
fades between states, a show.

## See also

[Lights & shows](../concepts/lights-and-shows.md) · [Show player](show-player.md)
