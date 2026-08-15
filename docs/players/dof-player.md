# DOF player

Fires DOF (DirectOutput) events for physical feedback — solenoids, shakers, beacons, button
lights.

```vbscript
With .DOFPlayer()
    With .EventName("s_Bumper1_active")
        .DOFEvent = 110
        .Action = "DOF_PULSE"
    End With
End With
```

| Property | Meaning |
|---|---|
| `DOFEvent` | DOF event number from your table's DOF config |
| `Action` | `DOF_ON`, `DOF_OFF`, or `DOF_PULSE` |

## Typical use

```vbscript
With .DOFPlayer()
    ' momentary
    With .EventName("s_LeftSlingshot_active")
        .DOFEvent = 101
        .Action = "DOF_PULSE"
    End With

    ' held for the duration of a mode
    With .EventName("mode_multiball_started")
        .DOFEvent = 140
        .Action = "DOF_ON"
    End With
    With .EventName("mode_multiball_stopping")
        .DOFEvent = 140
        .Action = "DOF_OFF"
    End With
End With
```

`DOF_ON` stays on until something turns it off — pair it with a stop entry, or it persists past the
mode.

## From a show step

DOF can also be fired from inside a show, which keeps an effect synchronised with its lighting:

```vbscript
With .AddStep(Null, Null, 0.2)
    .Lights = Array("(lights)|100|ff0000")
    With .DOFEvent(110)
        .Action = "DOF_PULSE"
    End With
End With
```

## Notes

- Event numbers come from your table's DOF configuration; GLF passes them through unchanged.
- Tables without DOF hardware are unaffected — the calls are no-ops.

## See also

[Lights & shows](../concepts/lights-and-shows.md)
