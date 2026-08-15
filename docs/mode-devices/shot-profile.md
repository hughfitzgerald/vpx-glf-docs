# Shot profile

A profile defines the states a [shot](shot.md) moves through and how each looks. Profiles are
shared: one `off_on_color` profile serves every insert on the table, with the shot supplying the
light and colour through tokens.

```vbscript
With GlfShotProfiles("off_on_color")
    With .States("unlit")
        .Show = "off"
        .Key = "key_off_on_color_unlit"
    End With
    With .States("on")
        .Show = "led_color"
        .Key = "key_off_on_color_on"
    End With
End With
```

Declare shared profiles in `ConfigureGlfDevices` with the global `GlfShotProfiles(name)`, before
any mode that uses them. A profile needed by only one mode can be declared on the mode instead:
`With .ShotProfiles("name")`.

## States

States are ordered by declaration: the first is index 0, which is where shots start. Each state
is configured like a [show player](../players/show-player.md) entry:

| Property | Default | Meaning |
|---|---|---|
| `Show` | — | Show to play while in this state |
| `Key` | `""` | Unique handle — **required in practice** |
| `Speed` | 1 | Divides step durations |
| `Loops` | -1 | -1 forever, 0 once, N repeats |
| `Priority` | 0 | Added to the shot's mode priority |
| `Tokens` | — | Default token values, overridden by the shot's |
| `EventsWhenCompleted` | — | Events when a non-looping show ends |

Give every state a distinct `Key`. Keys identify the show on the light stack; duplicates across
states cause one state's lighting to interfere with another's.

## Profile-level properties

| Property | Default | Meaning |
|---|---|---|
| `AdvanceOnHit` | `True` | Whether a hit advances the state |
| `Block` | `False` | Block by shot name rather than shot-plus-profile |
| `StateNamesToRotate` | all | States eligible for [shot group](shot-group.md) rotation |
| `StateNamesNotToRotate` | none | States excluded from rotation |

`AdvanceOnHit = False` gives a shot that reports hits but only changes state when told to via
control events — useful when a shot's appearance is driven by mode logic rather than by being hit.

## Tokens

Shows in a profile use `(token)` placeholders; the shot fills them:

```vbscript
With GlfShotProfiles("flicker_on")
    With .States("unlit")
        .Show = "off"
        .Key = "key_flicker_on_unlit"
    End With
    With .States("on")
        .Show = "flicker_color_on"
        .Key = "key_flicker_on_on"
        .Speed = 4
    End With
End With
```

```vbscript
With .Shots("left_ramp")
    .Profile = "flicker_on"
    .Switch = "s_LeftRamp"
    With .Tokens()
        .Add "lights", "L12"
        .Add "color", "8800ff"
    End With
End With
```

A state can also supply its own tokens, which the shot's override. That's how a profile pins one
light while leaving colour to the shot:

```vbscript
With GlfShotProfiles("shoot_again")
    With .States("unlit")
        .Show = "off"
        .Key = "key_shoot_again_unlit"
        With .Tokens()
            .Add "lights", "L03"
        End With
    End With
    With .States("flashing")
        .Show = "flash_color_with_fade"
        .Key = "key_shoot_again_flashing"
        .Speed = 2
        .Priority = 5000
        With .Tokens()
            .Add "lights", "L03"
            .Add "fade", 500
        End With
    End With
End With
```

## Three-state example

```vbscript
With GlfShotProfiles("lit_flashing_collected")
    With .States("unlit")
        .Show = "off"
        .Key = "key_lfc_unlit"
    End With
    With .States("lit")
        .Show = "led_color"
        .Key = "key_lfc_lit"
    End With
    With .States("flashing")
        .Show = "flash_color"
        .Key = "key_lfc_flashing"
        .Speed = 4
    End With
End With
```

Hitting a shot on this profile takes it unlit → lit → flashing, and it stops there. State names
appear in the events the shot posts (`myshot_lit_hit`), so name them for what they mean in the
rules, not for what they look like.

## Notes

- Shows referenced by a profile must exist before the profile is declared.
- Profile names are global. A second `GlfShotProfiles("default")` returns the existing built-in
  rather than a new one.
- The built-in `default` profile is two states, `on` (flash) and `off`.

## See also

[Shot](shot.md) · [Lights & shows](../concepts/lights-and-shows.md)
