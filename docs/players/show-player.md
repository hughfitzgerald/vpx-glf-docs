# Show player

Plays and stops [light shows](../concepts/lights-and-shows.md) in response to events.

```vbscript
With .ShowPlayer()
    With .EventName("jackpot_lit")
        .Key = "key_jackpot_show"
        .Show = "flash_color"
        .Speed = 4
        .Loops = -1
        .Priority = 500
        With .Tokens()
            .Add "lights", "L20"
            .Add "color", JackpotColor
        End With
    End With
End With
```

## Properties

| Property | Default | Meaning |
|---|---|---|
| `Show` | — | Name of a show created with `CreateGlfShow` |
| `Key` | `""` | Handle used to stop it later — keep unique |
| `Action` | `"play"` | `play` or `stop` |
| `Speed` | 1 | Divides step durations; higher is faster |
| `Loops` | -1 | -1 forever, 0 once, N additional repeats |
| `Priority` | 0 | Added to the mode's priority |
| `Tokens` | — | Values for `(token)` placeholders in the show |
| `EventsWhenCompleted` | — | Events posted when a non-looping show finishes |
| `SyncMs` | 0 | Aligns the show's start to a time grid |
| `BlockQueue` | `False` | Hold a queue event until the show completes |

## Stopping

Quote the key you started with:

```vbscript
With .EventName("jackpot_collected")
    .Key = "key_jackpot_show"
    .Action = "stop"
End With
```

Looping shows need this. One-shot shows (`Loops = 0`) end themselves, and everything a mode started
is cleaned up when it stops — but an explicit stop on `mode_<name>_stopping` is the habit that
avoids surprises.

## Tokens

Tokens are what make one show serve twenty inserts:

```vbscript
With .EventName("play_bumper1_show")
    .Key = "key_bumper1_show"
    .Show = "flash_color_with_fade"
    .Speed = 15
    .Loops = 0
    With .Tokens()
        .Add "lights", "tBL1"
        .Add "fade", 500
        .Add "color", Bumper1Color
    End With
End With
```

The `lights` token accepts a light name or a tag.

## Chaining

`EventsWhenCompleted` fires when a non-looping show ends, which is how you sequence animations:

```vbscript
With .EventName("start_intro")
    .Key = "key_intro"
    .Show = "wizard_intro"
    .Loops = 0
    .EventsWhenCompleted = Array("intro_finished")
End With
```

## Notes

- Shows must exist before the mode that references them is created. A missing show name fails
  silently.
- `Speed` divides durations: a 0.5 s step at speed 4 lasts 125 ms.
- Priority is *added to* the mode's, so a show at `Priority = 500` inside a priority-200 mode sits
  at 700 on the light stack.

## See also

[Lights & shows](../concepts/lights-and-shows.md) · [Light player](light-player.md) ·
[Shot profile](../mode-devices/shot-profile.md)
