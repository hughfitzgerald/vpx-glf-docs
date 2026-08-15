# Drop targets & stand-up targets

Both are declared as devices rather than added to `glf_switches` — GLF wires their switches itself.

**Neither drop targets nor stand-up targets belong in `glf_switches` or any other GLF collection.**

---

## Drop targets

```vbscript
With CreateGlfDroptarget("drop1")
    .Switch = "s_DT1"
    .KnockdownEvents = Array("DT1_knockdown")
    .ResetEvents = Array("ball_started", "reset_target_bank")
    .ActionCallback = "DT1Callback"
    .UseRothDroptarget = True
    .RothDTSwitchID = 1
End With
```

### Configuration

| Property | Default | Meaning |
|---|---|---|
| `Switch` | — | The target's switch |
| `ActionCallback` | — | Function that drives the target |
| `KnockdownEvents` | — | Events that knock the target down |
| `ResetEvents` | — | Events that raise it |
| `EnableKeepUpEvents` | — | Events that enable keep-up |
| `DisableKeepUpEvents` | — | Events that disable keep-up |
| `UseRothDroptarget` | `False` | Use Rothbauer drop target physics |
| `RothDTSwitchID` | -1 | Index into the Roth drop target array |
| `ExcludeFromBallSearch` | `False` | Skip during ball search |

### The action callback

Receives a number identifying the action:

| Value | Action |
|---|---|
| 0 | Reset (raise) |
| 1 | Knockdown |
| 3 | Enable keep-up |
| 4 | Disable keep-up |

```vbscript
Sub DT1Callback(action)
    Select Case action
        Case 0 : DTRaise 1
        Case 1 : DTDrop 1
    End Select
End Sub
```

### Roth drop targets

Most modern tables use the Rothbauer drop target system for physics. Set `UseRothDroptarget = True`
and `RothDTSwitchID` to the target's index in your Roth array; GLF then routes hits into `DTHit`
rather than posting switch events directly, and the Roth code reports state back.

### Events posted

| Event | When |
|---|---|
| `drop_target_<name>_down` | Target went down |
| `drop_target_<name>_up` | Target came up |

### Banks

A bank is a group of drop targets. Model each as a [shot](../mode-devices/shot.md) and collect them
in a [shot group](../mode-devices/shot-group.md), which gives you completion detection for free:

```vbscript
With .Shots("dt1")
    .Profile = "off_on_color"
    .Switch = "s_DT1"
    With .Tokens()
        .Add "lights", "L33"
        .Add "color", TargetBankColor
    End With
End With
' … dt2, dt3 …

With .ShotGroups("target_bank")
    .Shots = Array("dt1", "dt2", "dt3")
End With

With .EventPlayer()
    .Add "target_bank_on_complete", Array("bank_complete_award", "reset_target_bank")
End With
```

---

## Stand-up targets

Simpler — no coil, so no callback:

```vbscript
With CreateGlfStanduptarget("target1")
    .Switch = "s_ST1"
    .UseRothStanduptarget = True
    .RothSTSwitchID = 1
End With
```

| Property | Default | Meaning |
|---|---|---|
| `Switch` | — | The target's switch |
| `UseRothStanduptarget` | `False` | Use Rothbauer stand-up target physics |
| `RothSTSwitchID` | -1 | Index into the Roth stand-up array |

A stand-up target posts the ordinary switch events, so you use it like any other switch:

```vbscript
With .EventPlayer()
    .Add "s_ST1_active", Array("standup_hit", "score_5000")
End With
```

Declaring it as a device rather than adding it to `glf_switches` is what lets GLF route hits
through the Roth system when enabled.

## See also

[Shot](../mode-devices/shot.md) · [Shot group](../mode-devices/shot-group.md)
